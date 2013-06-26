#!/usr/bin/env python

import logging
import win32api
import win32gui
import win32con
import win32process
import getopt
import sys
import datetime
import time
import pywintypes
import os

def usage():
    progname = os.path.basename(sys.argv[0])
    print """%s - Zap Windows by titles (WM_CLOSE version)
    
    Zap Windows by given titles, sending them WM_CLOSE message.
    Not foolproof as it can hang, if the process in the window does not
    respond to the message.

    Usage:

    %s [-l][-o secs][-s secs] title [title ... title]

    Options:

       -l 
          Execute in the indefinite loop. Without this option, program
          executes once then exits.

       -o SECS
          Close windows, whose owner process is older than SECS seconds.
          This is optional, when not specified, it unconditionally closes 
          windows with given titles.

       -s SECS 
           In loop mode (option -l), number of seconds to sleep between
           each loop. This is optional, when not specified, 60 seconds is
           assumed.
        
       -h 
           Print this message.

    Parameters:

        title
           Window title.
           Note: Enclose each window title with " like:
             "title1" "title2" "title3"
             
    Example:
    
        # Execute in the loop and close each cmd.exe window if its 
        # owning process is older than half an hour
        
        python zapwin.py -o 1800 -l "C:\WINNT\system32\cmd.exe"
    """ % (progname, progname)
           

def killwin(whandle):
    logging.info('.. Sending WM_CLOSE message')
    # win32api.SendMessage(whandle, win32con.WM_CLOSE, -1)
    try:
        win32gui.SendMessageTimeout(whandle, win32con.WM_CLOSE, -1, -1, 
         win32con.SMTO_NORMAL, 5000)
    except Exception, e:
        logging.warning('/E/: %s' % e)

def _not_used_killwin():
    procids = win32process.GetWindowThreadProcessId(whandle)
    if not procids:
        logging.error('No processes!')
        return
    pid = procids[1]
    logging.info('PID: %s' % pid)
    phandle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, 0, pid)
    win32api.TerminateProcess(phandle, 0)
    win32api.CloseHandle(phandle)


def getOpenTime(whandle):
    procids = win32process.GetWindowThreadProcessId(whandle)
    if not procids:
        logging.error('No processes')
        return None
    pid = procids[1]
    phandle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION, 0, pid)
    if not phandle:
        logging.error('Cannot open process with pid=%s' % pid)
        return None
    times = win32process.GetProcessTimes(phandle)
    if not times:
        raise RuntimeError('GetProcessTimes did not return anything!')
    ct =  times['CreationTime']
    return ct
        
def zapwin(titles, opentime = 0):
    for title in titles:
        whandle = win32gui.FindWindow(None, title)
        if not whandle:
            logging.info('Window "%s" not found' % title)
            continue
        logging.info('Window: %s' % title) 
        if not opentime:
            killwin(whandle) 
        else:
            ct = getOpenTime(whandle)
            now = pywintypes.Time(time.gmtime())
            td = int(now) - int(ct)
            if td > opentime:
                logging.info('..Sending WM_CLOSE, process runs for %s secs' %
                 td)
                killwin(whandle)
            else:
                logging.info('..Process only runs for %s secs, not closing' %
                 td)

def main():
    datefmt = "%d/%m/%Y %H:%M:%S"
    lformat = "%(asctime)s %(levelname)-8s %(funcName)s: %(message)s"
    opentime = 0
    sleeptime = 60
    loop = False
    titles = []

    opts, args = getopt.getopt(sys.argv[1:], 'hlo:s:')
    for opt, arg in opts:
        if opt == '-l':
            loop = True
        if opt == '-o':
            try:
                opentime = int(arg)
                if opentime < 0:
                    raise RuntimeError(
                     'Parameter of -o option cannot be negative!')
            except Exception, e:
                logging.error('Error while processing -o option (%s)' % e)
                sys.exit(1)
        if opt == '-s':
            try:
                sleeptime = int(arg)
                if sleeptime <= 0:
                    raise RuntimeError(
                     'Parameter of -s option cannot be lower than 0')
            except Exception, e:
                logging.error('Error while processing -s option (%s)' % e)
                sys.exit(1)
        if opt == '-h':
            usage()
            sys.exit(0)
    titles = args
    if not titles:
        logging.error('Illegal number of arguments, missing window titles')
        usage()
        sys.exit(2)

    try:
        logging.basicConfig(level = logging.INFO, format = lformat,
         datefmt = datefmt)
        while True:
            zapwin(titles, opentime)
            if not loop:
                break
            logging.info('Waiting for %s secs..' % sleeptime)
            time.sleep(sleeptime)
            logging.info('Lets go')
    except KeyboardInterrupt:
        logging.info('INTERRUPT, BYE!')
    except Exception, e:
        logging.exception('/E/ Caught exception (%s)' % e)
        sys.exit(3)
    sys.exit(0)

if __name__ == '__main__':
     main()
