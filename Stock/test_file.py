import sys,os
from datetime import datetime

now = datetime.now()
file = open("data/log_1/%s-%s-%s.txt" %  (now.year,now.hour,now.day),'a+')
#os.path.exists("data/log_1/%s-%s-%s.txt" % (now.month,now.hour,now.day))
