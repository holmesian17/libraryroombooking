import openpyxl, sys
wb = openpyxl.load_workbook('Library Room Booking.xlsx')

def room_sheet():
  if room == 'p':
    sheet = wb.get_sheet_by_name('Pye')
  elif room == 's':
    sheet = wb.get_sheet_by_name('Scriver') 
def equip_initials():
  print('Do they need equipment?')
  equip_need = input()
  equip_need = equip_need.lower()
  if equip_need.startswith('n'):    #TODO: equip column - if n then "none"
    print("Are there any notes you'd like to add?") #TODO: if n then "none"
    noteQuestion = input()
    noteQuestion = noteQuestion.lower()
    if noteQuestion.startswith('y'):
      print('Enter your note:')
      note = input()
      sheet['F1'] = note
      print("Enter the staff member's initials")
      initials = input()
      initials = initials.lower()
      sheet['G1'] = initials
    else:  
      print("Enter the staff member's initials")
      initials = input()
      initials = initials.lower()
      sheet['G1'] = initials
  elif equip_need.startswith('y'):
    print('What equipment?')
    equip = input()
    print("Are there any notes you'd like to add?")
    noteQuestion = input()
    noteQuestion = noteQuestion.lower()
    if noteQuestion.startswith('y'):
      print('Enter your note:')
      note = input()
      print("Enter the staff member's initials")
      initials = input()
      while True:
        break
    else:
      print("Enter the staff member's initials")
      initials = input()
      while True:
        break
def ask():    
  while True:
    print('What room is being booked? (P, S)')
    room = input()
    room = room.lower()
    if room == 'p':
      sheet = wb.get_sheet_by_name('Pye')
      sheet['B1'] = 'Pye'
    elif room == 's':
      sheet = wb.get_sheet_by_name('Scriver')
      sheet['B1'] = 'Scriver'
    else:
      continue    
    print('When is the booking date? (mm/dd/yyyy)')
    date = input()
    sheet['C1'] = date
    print('What time is the booking? (24hrs - 4 digits)')
    later_time = input()
    sheet['D1'] = later_time
    print('What is the borrower barcode?')
    barcode = input()
    print(barcode)
    print('Is this correct? (Y or N)')
    verify = input()
    verify = verify.lower()
    while True:
      if verify.startswith('n'):
        continue
      else:
        sheet['A1'] = barcode
        print('How many people will be using the room?')
        persons = input()
        equip_initials()
        wb.save('Library Room Booking.xlsx')
        sys.exit("Done")
        break
ask()







#look at Chase's code, perhaps put the responses in a dictionary and then use that to
	#export? 
#need to file it into a dictionary or database - something where you can call any value 
#just send the information to an excel spreadsheet?
#need to record the date and time the booking is taken
#have it automatically add it into outlook meeting room calendar
#be able to pull the statistics for which room, walk-in or no, date and time, etc.
#have an alert come up 2 hours after their booking to have the staff person check with them
#how about pulling the information from the database? gathering the statistics?
#booking into the future? how to add in the date and time - do we want an alarm for when a booking is supposed to be occurring?
#how to keep someone from booking a room during a time we are not open?
#making an alarm that will go off after their 2 hours is up - asks if wanting to continue
#if B then go through a different set of steps or questions for Bunday room?
  #would this then automatically get sent to Teresa?
#How would someone edit a booking? Would they just have to maunally change things in the excel sheet/calendar?
