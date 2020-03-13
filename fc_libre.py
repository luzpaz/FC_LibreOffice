'''
March 11,20 (SP) Succesfully launch soffice and connect to service on port 2002

'''
import socket
import subprocess

import uno

def check_socket(host, port):
    # Test if port up and listening
    s = socket.socket()
    try:
        s.connect((host, port))
        return True
    except socket.error as e:
        return False
    finally:
        s.close()


result = check_socket('127.0.0.1', 2002)
if result == False:
    # Start the LibreOffice service on port 2002
    run_soffice = [
        'soffice',
        '--accept=socket,host=localhost,port=2002;urp;StarOffice.Service',
        '--nologo',
    ]
    subprocess.Popen(run_soffice)

# Create connection to LibreOffice on local ocmputer listening on port 2002
localContext = uno.getComponentContext()
resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
doc = desktop.getCurrentComponent()


