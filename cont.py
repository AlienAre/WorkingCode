import win32com.client

o = win32com.client.gencache.EnsureDispatch("Outlook.Application")
ns = o.GetNamespace("MAPI")

adrLi = ns.AddressLists.Item("Global Address List")
contacts = adrLi.AddressEntries
numEntries = adrLi.AddressEntries.Count
print numEntries
for i in contacts:
	name = i.Name
	print name
	
	