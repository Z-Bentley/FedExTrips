import time

days = ['Sunday', 'Monday', 'Tuesday', 
        'Wednesday', 'Thursday', 'Friday', 'Saturday']

def dayCheck(days):
    for day in days:
        time.sleep(1)
        if day == 'Saturday':
            print("Yay it\'s Saturday, Nicole!")
        else:
            print('Darn today is %s, if only it were Saturday.' % (day))

dayCheck(days)