from datetime import datetime, timedelta

t = datetime(2011,10,15,12,32,15)

print(t)
print(t.minute)
# round down
diff = t.minute % 30
t_sub = timedelta(minutes=diff)
print(t_sub)
t_down = t - t_sub
print(t_down)

# round up
diff = t.minute % 30
t_sub = timedelta(minutes=30-diff)
print(t_sub)
t_up = t + t_sub
print(t_up)
