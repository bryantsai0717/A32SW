import datetime

def round_minutes(dt, direction, resolution):
    print(dt.minute)
    print(resolution + (1 if direction == 'up' else 0))
    print((dt.minute // resolution + (1 if direction == 'up' else 0)))
    new_minute = (dt.minute // resolution + (1 if direction == 'up' else 0)) * resolution
    print(new_minute)
    print(dt)
    print(new_minute - dt.minute)
    return dt + datetime.timedelta(minutes=new_minute - dt.minute)

#for hour, minute, resolution in ((17, 34, 30), (12, 58, 15), (14, 1, 60)):
hour = 17
minute = 34
resolution = 30
dt = datetime.datetime(2014, 8, 31, hour, minute)
for direction in 'up', 'down':
    print('{} with resolution {} rounded {:4} is {}'.format(dt, resolution, direction,
        round_minutes(dt, direction, resolution)))
