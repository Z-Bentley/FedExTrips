# not useable since outlook updated
import win32com.client as win

# To calculate the time variance of scheduled arival to actual arival
print('check 0')
ol = win.Dispatch('Outlook.Application')
print('check 1')

olMainItem = 0x0
print('check 2')

newMail = ol.CreateItem(olMainItem)
print('check 3')

newMail.Subject = 'Testing'
print('check 4')

newMail.To = 'jksebastian@fedex.com'
newMail.CC = 'zachary.bentley@fedex.com'

newMail.Body = 'I am testing something, so why not RUN IT!'
print('check 5')

newMail.Display()
print('last check')
newMail.Send()