import os
from sympy import matrix_multiply_elementwise
import win32com.client as win32

olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

mailItem = olApp.CreateItem(0)
mailItem.Subject = 'Prueba de envío de correo con Python'
mailItem.BodyFormat = 1
mailItem.Body = 'Prueba de funcionamiento correcta'
mailItem.To = 'juan.nieto@exsis.com.co'
mailItem.CC = 'nicolas.mendoza@exsis.com.co ; joseph.rojas@exsis.com.co'
mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('nicolas.mendoza@exsis.com.co')))

mailItem.Attachments.Add(os.path.join(os.getcwd(),'2022-03-24-Pool-2034.0-Desarrollador AS400.xlsx'))

mailItem.Display()
mailItem.Save()
mailItem.Send()
