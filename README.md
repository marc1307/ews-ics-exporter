# ews-ics-exporter
Exports Microsoft Exchange Calendar to a .ics file

## Install
```
~/ews-ics-exporter$ python3 -m venv env
~/ews-ics-exporter$ source env/bin/activate
~/ews-ics-exporter$ pip install -r requirements.txt
```

## Run
```
~/ews-ics-exporter$ ews-ics-exporter/env/bin/python ews-ics-exporter/export.py
```

## Cronjob
```
*/15	8-17	*	*	1-5	ews-ics-exporter/env/bin/python ews-ics-exporter/export.py > /dev/null
0	7,18,20,22	*	*	1-5	ews-ics-exporter/env/bin/python ews-ics-exporter/export.py > /dev/null
0	8,12,16,20	*	*	0,6	ews-ics-exporter/env/bin/python ews-ics-exporter/export.py > /dev/null
```