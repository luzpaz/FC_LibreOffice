
import uno
import os
import subprocess
import time
import shlex

import FreeCAD as App
import FreeCADGui as Gui




def libreoffice_init():
    #command_line = '/usr/bin/soffice --calc --nologo --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager" "/home/polifemos/Documents/fc_spreadsheet.ods" '
    command_line = '/usr/bin/soffice --nologo --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager" '
    args = shlex.split(command_line)
    p = subprocess.Popen(args)

    # wait for LibreOffice to initialize
    time.sleep(2)

    ## code to initialize OfficeLibre ##
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localContext)

    # connect to the running office
    ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    model = desktop.getCurrentComponent()

def libreoffice_open_doc(docname = None):
    print(docname)


libreoffice_init()
libreoffice_open_doc("/home/polifemos/Documents/fc_spreadsheet.ods")


#TODO User file dialog


#TODO get list of tabs in spreadsheet

#TODO read "named ranges" from spreadsheet

#TODO query FreeCAD object fields 


