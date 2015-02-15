#!/usr/bin/env python
#-------------------------------------------------------------------------------
# Name:        launch_manager 
# Purpose:     To make sure the correct Windows display configuration are True
#              before the launching of the GSDM and GSDMC.
#
# Author:      Sean Wiseman
#
# Created:     09/05/2014
#-------------------------------------------------------------------------------

import os, sys
import time
import datetime
import win32api
import win32com
import win32com.client
import pythoncom
import subprocess
import platform
import threading
import Queue 
import shelve
from ast import literal_eval            # from converting literal str to dict
sys.path.append('Launch Manager Utils')
import wyXML
from wyLogger import WyLogger
from wyNotify import WyNotify

class Launcher(object):
    def __init__(self, xml_conf):
        """
        Arguments:
            xml_conf  -- configuration file object
            
        """
        self.xml_conf = xml_conf                                                # config file object
        self.xlocal = os.path.dirname(__file__)                                 # local dir path
        self.os_type = self.osCheck()                                           # checks os type
        self.mDrive = self.dirCheck()                                           # main drive letter
        self.gsdm_path = os.path.join(self.mDrive, "ThinClient")                # gsdm path
        self.gsdmc_path = os.path.join(self.mDrive, "GSDMController")           # gsdmc path
        self.log_file = WyLogger('Launch_manager_log', weekRotate=True, startDayValue=0)
        
        try:
            self.config = wyXML.WyXML('LaunchManagerConfig.xml')
        except Exception, e:
            self.log_file.logEntry('{0}\nUnable to load configuration file'.format(e))
            
        self.startup_flag = True                                                # defines if process is from startup
        self.metrics_match = False                                              # used to determine if metrics are correct
        self.retry_limit = int(self.xml_conf.find('retryLimit'))                # number of times the launcher will retry after start up failure
        self.retry_counter = 0                                                  # keeps track of how many minutes between retry_count resets
        self.startup_delay = int(self.xml_conf.find('startupDelay'))            # start up delay (after logging on)
        self.retry_forget_time = int(self.xml_conf.find('retryForgetTime'))     # amount of time before the retry limit is reset (in minutes)
        self.forget_counter = 0                                                 # used to keep track of time
        self.redetect_limit = int(self.xml_conf.find('redetectLimit'))          # amount of times self.reDetectMonitors() wil be run before reboot
        
        self.sleep_active, self.sleep_start, self.sleep_stop = self.sleepTimeManager() # Set sleep period and times for the GSDM
        
        self.current_time = time.strftime('%H:%M')                              # keeps track of current time
        self.system_awake = True                                                # marked as True whilst the system is out of sleep period
        self.slave_unit = self.str2bool(self.xml_conf.find('slaveUnit'))        # Set whether an MVS unit is a slave (if True GSDMC will never start)
        self.logQ = Queue.Queue()                                               # queue to process log file entries
        self.timeQ = Queue.Queue()                                              # queue to process sleep functionality
        self.log_lock = threading.Lock()                                        # lock for log entry access
        
        # Notification gui
        self.notify_gui = WyNotify('Launch Manager')
        self.notify_hidden = False      
        
        # Threads ----------------------------------------------------------------
        self.log_thread = threading.Thread(target=self.logManager, args=())
        self.notify_thread = threading.Thread(target=self.notify_gui.run, args=())
        self.log_thread.daemon = True
        self.notify_thread.daemon = True
        # ------------------------------------------------------------------------
        # test modification for git hub
         
    def run(self):
        """ main function """
        self.log_thread.start()        # start logQ thread
        self.notify_thread.start()     # start notify thread
        self.notifyPut('Starting up...')
        self.logQ.put('Start up ***************************')
        
        tick = self.startup_delay
        for x in range(self.startup_delay):
            self.notifyPut('Starting Launch Process in {0} seconds'.format(tick))
            time.sleep(1)
            tick -= 1
        # First check on start up ------------------------
        self.getSystemAwake()
        self.getRetryCount()
        self.getSavedMetrics()
        self.reDetectMonitors()
        time.sleep(2)
        
        print 'Before startup test system_awake = {0}'.format(self.system_awake) # TESTING ++++++++++++++++
        
        if self.sleepPeriodValidate():
            self.system_awake = False
            if self.slave_unit == False:
                self.startGsdmc()
                self.startup_flag = False
        
        if self.system_awake:
            if self.metrics_match == False:
                self.noMatchLaunch()
            else:
                self.matchLaunch()
        elif not self.system_awake:
            self.notifyPut('GSDM is currently in a Sleep period')
            time.sleep(2)
            self.logQ.put('Sleep period: Start = {0} -- Stop = {1}'.format(self.sleep_start, self.sleep_stop))
            time.sleep(2)
            self.notifyPut('Current time is {0} -- Now entering Sleep period of the GSDM'.format(self.current_time))
            self.logQ.put('Current time is {0} -- Now entering Sleep period of the GSDM'.format(self.current_time))
            time.sleep(2)
            self.notifyPut('Sleep period will Stop at {1}'.format(self.sleep_start, self.sleep_stop))
        # ------------------------------------------------
        try:
            while True:
                    time.sleep(60)
                    self.retryManager()                             # manage retry resets
                    self.current_time = time.strftime('%H:%M')      # update current time
                    self.sleepCheck()                               # check sleep state and action
                    
        except Exception, e:
            self.logQ.put('{0} - Main run function failed'.format(e))
        print('EXIT: should never see this!')           #TESTING ++++++++++++++++++++++++++++++++++++
        self.logQ.put('EXIT: should never see this!')   #TESTING ++++++++++++++++++++++++++++++++++++
        
        # ------------------------------------------------
    
    def retryManager(self):
        """ keeps track of retry_count retry_forget_time and resets if reached """
        if self.retry_counter >= self.retry_forget_time:
            self.retry_counter = 0
            if self.retry_count > 0:
                self.retry_count = 0
                self.db = shelve.open(os.path.join(self.xlocal, 'Launch Manager Utils\\launch.data'))
                self.db['retry_count'] = self.retry_count
                self.db.close()
                self.logQ.put('Retry count of {0} has been reset to {1} after Retry forget time of {2}'.format(
                        self.retry_limit,
                        self.retry_count,
                        self.retry_forget_time))
        self.retry_counter += 1
        
    def sleepCheck(self):
        """ check sleep state and action based on current time """
        if self.sleep_active:
            if self.system_awake:
                if self.sleepPeriodValidate():
                    self.logQ.put('Sleep period: Start = {0} -- Stop = {1}'.format(self.sleep_start, self.sleep_stop))
                    self.notifyPut('Current time is {0} -- Now entering Sleep period of the GSDM'.format(self.current_time))
                    self.logQ.put('Current time is {0} -- Now entering Sleep period of the GSDM'.format(self.current_time))
                    time.sleep(3)
                    self.stopGsdm()
                    time.sleep(8)
                    
                    self.system_awake = False
                    self.db = shelve.open(os.path.join(self.xlocal, 'Launch Manager Utils\\launch.data'))
                    self.db['system_awake'] = self.system_awake
                    print 'After setting to false system_awake = {0}'.format(self.system_awake) # TESTING ++++++++++++++++
                    print 'After setting to false db[system_awake] = {0}'.format(self.db['system_awake']) # TESTING ++++++++++++++++
                    self.db.close()
                    
                    self.notifyPut('Sleep period will Stop at {1}'.format(self.sleep_start, self.sleep_stop))
                else:
                    pass
            elif not self.system_awake:
                if not self.sleepPeriodValidate():
                    
                    self.system_awake = True
                    self.db = shelve.open(os.path.join(self.xlocal, 'Launch Manager Utils\\launch.data'))
                    self.db['system_awake'] = self.system_awake
                    print 'After setting to true system_awake = {0}'.format(self.system_awake) # TESTING ++++++++++++++++
                    print 'After setting to true db[system_awake] = {0}'.format(self.db['system_awake']) # TESTING ++++++++++++++++
                    self.db.close()
                    
                    self.notifyPut('Sleep period: Start = {0} -- Stop = {1}'.format(self.sleep_start, self.sleep_stop))
                    self.logQ.put('Sleep period: Start = {0} -- Stop = {1}'.format(self.sleep_start, self.sleep_stop))
                    time.sleep(3)
                    self.notifyPut('Current time is {0} -- Now exiting Sleep period of the GSDM'.format(self.current_time))
                    self.logQ.put('Current time is {0} -- Now exiting Sleep period of the GSDM'.format(self.current_time))
                    time.sleep(3)
                    self.checkMetrics()
                    time.sleep(1)
                    self.stopGsdm() # to clear any unwanted sessions
                    time.sleep(5)
                    if self.metrics_match == False:
                        self.noMatchLaunch()
                    else:
                        self.matchLaunch()
                else:
                    pass
        
    def sleepTimeManager(self):
        """ manages daily sleep time settings """
        days = {0:'monday', 1:'tuesday', 2:'wednesday', 3:'thursday', 4:'friday', 5:'saturday', 6:'sunday'}
        sleep_active = self.str2bool(self.xml_conf.find('active'))
        
        print 'sleep_active = {0}, is of type = {1}'.format(sleep_active, type(sleep_active))   #TESTING ++++++++++++++++++++++++++++++++++++
        
        start_time = str(self.xml_conf.find_attrib(days[datetime.datetime.today().weekday()], 'start'))
        stop_time = str(self.xml_conf.find_attrib(days[datetime.datetime.today().weekday()], 'stop'))
        
        print 'start_time = {0}, is of type = {1}'.format(start_time, type(start_time))         #TESTING ++++++++++++++++++++++++++++++++++++
        print 'stop_time = {0}, is of type = {1}'.format(stop_time, type(stop_time))            #TESTING ++++++++++++++++++++++++++++++++++++
        
        return sleep_active, start_time, stop_time
    
    def sleepPeriodValidate(self):
        """ validate whether the system is in its sleep period """
        # sleep_validate = False (not in sleep period)
        # sleep_validate = True (in sleep period)
        
        sleep_validate = None
        pre_midnight = '23:59'
        midnight = '00:00'
        
        # check if out of sleep period
        if self.current_time >= self.sleep_stop and self.current_time < self.sleep_start:
            sleep_validate = False
            
        # check if in sleep period
        elif self.current_time >= self.sleep_start and self.current_time <= pre_midnight:
            sleep_validate = True  
        elif self.current_time < self.sleep_stop and self.current_time > midnight:
            sleep_validate = True
            
        return sleep_validate
        
    def notifyPut(self, data):
        """ send text data to be displayed in the Notify gui """
        if self.notify_hidden:
            self.notify_gui.dataQ.put('*SHOW*')     # if gui is hidden when you send display data it will wake
            self.notify_hidden = False
        self.notify_gui.dataQ.put(data)
        
    def matchLaunch(self):
        """ Launch function when Metrics match """
        self.logQ.put('Success: Display Metrics match')
        #self.logQ.put('Starting GSDMC & GSDM... ')
        if self.startup_flag == True:
            if self.slave_unit == False:
                    self.startGsdmc()
                    self.startup_flag = False
        if self.system_awake:
            self.startGsdm()
    
    def noMatchLaunch(self):
        """ Launch function when Metrics don't match """
        self.retry_count += 1                       # keep count of retry attempts
        if self.retry_count <= self.retry_limit:
            for process in range(self.redetect_limit):  # only run for MAX detect limit
                time.sleep(3)
                self.checkMetrics()
                if self.metrics_match == False:
                    self.logQ.put('Display Metrics do not match, attempting to redetect')
                    self.reDetectMonitors()
                    time.sleep(2)
                else:
                    self.logQ.put('Success: Display Metrics match')
                    break
            if self.metrics_match == False:
                self.logQ.put('Display Metrics still do not match, restarting system')
                self.restartSystem()
            else:
                if self.startup_flag == True:
                    if self.slave_unit == False:
                        self.startGsdmc()
                        self.startup_flag = False
                self.startGsdm()
        else:
            self.logQ.put('Retry limit of {0} reached!'.format(self.retry_limit))
            self.logQ.put('Starting GSDMC & GSDM with current screen configuration... ')
            if self.startup_flag == True:
                if self.slave_unit == False:
                        self.startGsdmc()
                        self.startup_flag = False
            if self.system_awake:
                self.startGsdm()
    
    def logManager(self):
        """ takes args from logQ and applies to logfile entries """
        time.sleep(0.1)
        while True:
            try:
                time.sleep(0.2)
                data = self.logQ.get(block=False)
            except Queue.Empty:
                pass
            else:
                try:
                    self.log_lock.acquire() 
                    self.log_file.logEntry(data)
                    time.sleep(0.1)
                    self.log_lock.release()
                except:
                    print '*Unable to write to log file*'
    
    def checkMetrics(self):
        """ Compare current Metrics with saved Metrics """  
        # get current metrics
        self.notifyPut('Comparing Display Metrics')
        self.current_metrics = self.getCurrentMetrics()
        self.logQ.put('\nsaved metrics   = {0}  \ncurrent metrics = {1}'.format(self.saved_metrics, self.current_metrics))
        if len(self.current_metrics) == len(self.saved_metrics):
            self.metrics_match = True
            return
        else:
            self.metrics_match = False
    
    def getCurrentMetrics(self):
        """ Get current Metrics from OS """
        self.notifyPut('Obtaining Current Display Metrics')
        try:
            data = []
            data = win32api.EnumDisplayMonitors(None, None)
            screens = {}
            scrNum = 0
            for screen in data:
                screens[scrNum] = screen[2]
                scrNum += 1
            return screens 
        except Exception, e:
            self.logQ.put('{0} - Unable to capture current metrics'.format(e))
    
    def getSavedMetrics(self):
        """ Get saved metrics from gsdm config """
        # default metrics for a last resort
        self.default_metrics = {0: (0, 0, 1920, 1080), 1: (1920, 0, 3840, 1080), 2: (3840, 0, 5760, 1080), 3: (5760, 0, 7680, 1080)}
        # get saved metrics
        self.notifyPut('Obtaining Saved Display Metrics')
        try:
            gsdm_conf = wyXML.WyXML(os.path.join(self.gsdm_path, 'gsdm\\conf\\gsdm_cfg.xml'))
            if gsdm_conf.find('displayMetrics') != None:
                self.saved_metrics = literal_eval(gsdm_conf.find('displayMetrics'))
                #gsdm_conf.replace('displayMetrics', '')
                self.db = shelve.open(os.path.join(self.xlocal, 'Launch Manager Utils\\launch.data'))
                self.db['display_metrics'] = self.saved_metrics
                self.db.close()
            else:
                self.db = shelve.open(os.path.join(self.xlocal, 'Launch Manager Utils\\launch.data'))
                self.saved_metrics = self.db['display_metrics']
                self.db.close()
                
        except Exception, e:
            self.logQ.put('{0} - Unable to detect saved metrics from GSDM configuration'.format(e))
            self.saved_metrics = self.default_metrics
    
    def getRetryCount(self):
        """ Get retry count from launch_store.data or give default """
        try:
            self.db = shelve.open(os.path.join(self.xlocal, 'Launch Manager Utils\\launch.data'))
            if self.db['retry_count']:
                self.retry_count = self.db['retry_count']
                self.db.close()
            else:
                self.retry_count = 0
                self.db['retry_count'] = self.retry_count
                self.db.close()
                
        except Exception, e:
                self.log_file.logEntry('{0}\nUnable to load previous retry_count, setting value to 0'.format(e))
                self.retry_count = 0
                
    def getSystemAwake(self):
        """ Get system_awake from launch_store.data or give default """
        print 'start of getSystemAwak() system_awake = {0}'.format(self.system_awake) # TESTING ++++++++++++++++
        try:
            self.db = shelve.open(os.path.join(self.xlocal, 'Launch Manager Utils\\launch.data'))
            if self.db['system_awake'] == False:
                print 'start of if true - getSystemAwak() system_awake = {0}'.format(self.system_awake) # TESTING ++++++++++++++++
                self.system_awake = self.db['system_awake']
                self.db.close()
            else:
                self.system_awake = True
                self.db['system_awake'] = self.system_awake
                self.db.close()
                
            print 'End of getSystemAwak() system_awake = {0}'.format(self.system_awake) # TESTING ++++++++++++++++
                
        except Exception, e:
                self.log_file.logEntry('{0}\nUnable to load previous system_awake value, setting value to True'.format(e))
                self.system_awake = True
        
    def reDetectMonitors(self):
        """ Run Windows Detect on Display properties (WIN8 only!) """
        if self.os_type == 'Windows8':
            try:
                self.notifyPut('Running a quick monitor detect')
                self.checkMetrics()
                pythoncom.CoInitialize()                                            # Initialize COM lib on thread
                shell = win32com.client.Dispatch('WScript.Shell')
                time.sleep(0.1)
                subprocess.Popen(['control','desk.cpl'])
                time.sleep(2)
                shell.SendKeys("%C", 0)
                time.sleep(1)
                shell.SendKeys("%{F4}", 0)
            except Exception, e:
                self.logQ.put('{0} - Unable to redetect display(s)'.format(e))
        if self.os_type == 'WindowsXP':
            pass
            '''
            try:
                self.checkMetrics()
                expected_total = len(self.saved_metrics)                                
                missing_screens = len(self.saved_metrics) - len(self.current_metrics)
                
                if missing_screens > 0:
                    self.notifyPut('Attempting to redetect {0} missing monitors'.format(missing_screens))
                    pythoncom.CoInitialize()                                           # Initialize COM lib on thread
                    shell = win32com.client.Dispatch('WScript.Shell')
                    time.sleep(1)
                    subprocess.Popen(['control','desk.cpl'])
                    time.sleep(2)
                    shell.SendKeys("^+{TAB}", 0)
                    
                    time.sleep(0.5)
                    for x in range(expected_total + 1 - missing_screens , expected_total + 1):
                        time.sleep(0.5)
                        key = "{"+str(x)+"}"
                        time.sleep(0.1)
                        shell.SendKeys(key , 0)
                        for x in range(4):
                            time.sleep(0.5)
                            shell.SendKeys("{TAB}", 0)
                        time.sleep(0.5)
                        shell.SendKeys(" ", 0)
                        shell.SendKeys("%A", 0)
                        time.sleep(0.5)
                    shell.SendKeys("%{F4}", 0)
                    time.sleep(0.5)
                
            except Exception, e:
                self.logQ.put('{0} - Unable to redetect display(s)'.format(e))
                '''
        return
    
    def startGsdm(self):
        """ start the GSDM application """
        self.notifyPut('Starting the GSDM...')
        self.logQ.put('Starting the GSDM...')
        
        try:
            time.sleep(3)
            current_dir = os.getcwd()
            os.chdir(self.gsdm_path)
            os.startfile('gsdm_start.bat')
            self.logQ.put('GSDM started successfully')
            os.chdir(current_dir)
        except Exception, e:
            self.logQ.put('{0} - Unable to start the GSDM'.format(e))
        if not self.notify_hidden:
            self.notifyPut('*HIDE*')
            self.notify_hidden = True
        
    
    def stopGsdm(self):
        """ stop the GSDM application """
        self.notifyPut('Stopping the GSDM...')
        self.logQ.put('Stopping the GSDM...')
        try:
            current_dir = os.getcwd()
            os.chdir(self.gsdm_path)
            os.startfile('gsdm_stop.bat')
            self.logQ.put('GSDM stopped successfully')
            os.chdir(current_dir)
        except Exception, e:
            self.logQ.put('{0} - Unable to stop the GSDM'.format(e))
            
    def startGsdmc(self):
        """ start the GSDMC application """
        self.notifyPut('Starting the GSDMC...')
        self.logQ.put('Starting the GSDMC...')
        try:
            time.sleep(3)
            current_dir = os.getcwd()
            os.chdir(self.gsdmc_path)
            os.startfile('start.bat')
            self.logQ.put('GSDMC started successfully')
            os.chdir(current_dir)
        except Exception, e:
            self.logQ.put('{0} - Unable to start the GSDMC'.format(e))
        
    def restartSystem(self):
        """ exit thread and restart local machine """
        # save retry count between reboots
        try:
            self.notifyPut('Restarting System...')
            self.db = shelve.open(os.path.join(self.xlocal, 'Launch Manager Utils\\launch.data'))
            self.db['retry_count'] = self.retry_count
            self.db.close()
        except Exception, e:
            self.logQ.put('{0} - Unable to save retry count'.format(e))
            
        try:
            subprocess.call(['SHUTDOWN', '/f', '/r'])
        except Exception, e:
            self.logQ.put('{0} - Unable to restart Windows'.format(e))
        return
    
    def osCheck(self):
        '''
        Windows XP = WindowsXP
        Windows 7 = Windows7
        Windows 8 = Windowspost2008Server (converted to Windows8)
        '''
        osFound = platform.system() + platform.release()
        if osFound == 'Windowspost2008Server': # Clean up the return value of Windows 8
            osFound = 'Windows8'
        return osFound
    
    def dirCheck(self):
        # Detect the main install drive
        xDrive = None
        if os.path.exists('C:/ThinClient'):
            xDrive = 'C:\\'
        if os.path.exists('D:/ThinClient'):
            xDrive = 'D:\\'
        elif os.path.exists('C:/Receiver'):
            xDrive = 'C:\\'
        elif os.path.exists('D:/Receiver'):
            xDrive = 'D:\\'
        return xDrive
    
    def str2bool(self, val):
        """ allows conversion between strings and bools """
        return val.lower() in ('true','yes','t',1)

def main():
    launcher_config = wyXML.WyXML('LaunchManagerConfig.xml')
    test = Launcher(launcher_config)
    test.run()

if __name__ == '__main__':
    main()
    