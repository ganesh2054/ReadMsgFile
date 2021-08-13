# Read message of .msg file
### This module serves the following functions: 
- Read message of msg file, save message and attachments in your machine
- Send your message to recipient via outlook 

## create virtual environment named "myenv"
```
 python3 -m venv myenv           # rename environment as required
```
## activate the environment 
```
 source myenv/bin/activate  # for linux
 myenv\Scripts\activate     # for windows

```
## Install all requirements
```
pip install extract-msg
pip install exchangelib
```
## Call read_msg.py
```
$ python read_msg.py --h


positional arguments:
  -i INPUTFILE,           --inputFile
                             file from which message to be read
  -se SENDER EMAIL,       --s_email
                             outlook email of sender
  -p  PASSWORD,           --password
                             outlook password
  -re RECIPIENT EMAIL     --r_email
                             recipient email 
  -m  MESSAGE             --m message
                             your message for recipient               

optional arguments:
  -h, --help                show this help message and exit
```






