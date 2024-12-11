import os
import time
import subprocess
import win32com.client as win32Client

class SAPAuth:
    def __init__(self):
        self.SAP_BIN = "saplogon.exe"
        self.SAP_PROCESS = r"SAP Logon 800"   
        self.SAP_GUI_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\saplogon.exe"
        self.Rosenberger_auth_P80 = "Rosenberger Automotive Cabling (P80) Production"
        self.Rosenberger_auth_E80 = "Rosenberger Automotive Cabling (E80) Testsystem"
        self.pClient = ""
        self.pUsername = ""
        self.pPassword = ""
        self.Application = None
        self.Connection = None
        self.Session = None
        self.Root = None
        self.pSession = None
        self.SAPAlreadyOpened = False
        self.SAPSessionActive = False

    def GetLogin(self):
        if self.pSession is None:
            self.Login()
            self.pSession = self.Session
        return self.pSession

    def SetInstance(self, Username, Client):
        self.pUsername = Username
        self.pClient = Client

    def Login(self):
        print("Abriendo archivo SAP GUI..." + self.SAP_GUI_PATH)
        if not os.path.exists(self.SAP_GUI_PATH):
            print("El archivo no existe en la ruta especificada.")
            return

        self.SAPAlreadyOpened = self.IsProcessRunning(self.SAP_BIN)

        if not self.SAPAlreadyOpened:
            self.ExecuteAndWaitForSAP(self.SAP_BIN, self.SAP_GUI_PATH)
            self.pPassword = input("Ingresa tu contraseña: ")

        try:
            self.Root = win32Client.GetObject("SAPGUI")
        except Exception as e:
            print("No se pudo obtener la aplicación SAP GUI. Asegúrate de que SAP GUI esté abierto y que la opción de scripting esté habilitada.") 
            print(f"Error: {e}")
            return

        if self.Root is not None:
            self.Application = self.Root.GetScriptingEngine

            if self.Application.Children.Count > 0:
                self.Connection = self.Application.Children(0)
                self.SAPSessionActive = (self.Connection.Children.Count > 0)

                if self.SAPSessionActive:
                    self.Session = self.Connection.Children(0)
                else:
                    if not self.pPassword:
                        self.pPassword = input("Ingresa tu contraseña: ")
                    self.Connection = self.Application.OpenConnection(self.Rosenberger_auth_E80, True)
                    self.Session = self.Connection.Children(0)
                    if self.Session is not None:
                        self.LoginToSAP()
            else:
                if not self.pPassword:
                    self.pPassword = input("Ingresa tu contraseña: ")
                self.Connection = self.Application.OpenConnection(self.Rosenberger_auth_E80, True)
                self.Session = self.Connection.Children(0)
                if self.Session is not None:
                    self.LoginToSAP()

        if not self.Session:
            print("No se pudo obtener la sesión de SAP.")
            return

    def LoginToSAP(self):
        try:
            self.Session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = self.pClient
            self.Session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = self.pUsername
            self.Session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = self.pPassword
            self.Session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "EN"
            self.Session.findById("wnd[0]").sendVKey(0)
        except UnicodeDecodeError as e:
            print(f"Unicode decode error: {e}")
        except Exception as e:
            print(f"Error during SAP login: {e}")

    def IsProcessRunning(self, targetProcess):
        try:
            call = ['TASKLIST', '/FI', f'IMAGENAME eq {targetProcess}']
            output = subprocess.check_output(call, encoding='ansi')
            # Elimina encabezados de la salida y verifica si aparece el proceso
            processes = output.strip().split('\r\n')
            for process in processes:
                if targetProcess.lower() in process.lower():
                    return True
            return False
        except subprocess.CalledProcessError as e:
            print(f"Error ejecutando TASKLIST: {e}")
            return False


    def ExecuteAndWaitForSAP(self, SAP_BIN, SAP_GUI_PATH):
        if not os.path.exists(SAP_GUI_PATH):
            print(f"El archivo {SAP_GUI_PATH} no existe.")
            return
        print("Opening SAP GUI in 1, 2, 3...")
        os.system(f'"{SAP_GUI_PATH}"')
        while not self.IsProcessRunning(SAP_BIN):
            time.sleep(1)
            print("Waiting for SAP to open...")
        time.sleep(5)
        print("SAP opened successfully.")

# Example usage:
sap_auth = SAPAuth()
sap_auth.SetInstance("CAPECINA", "113")
session = sap_auth.GetLogin()