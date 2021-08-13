import  extract_msg
from exchangelib import Account, Credentials, DELEGATE, Message, Mailbox, \
  FileAttachment, HTMLBody
from exchangelib.items import SEND_ONLY_TO_ALL, SEND_ONLY_TO_CHANGED
from exchangelib.properties import DistinguishedFolderId
import argparse


def main():

  
  ap = argparse.ArgumentParser()

# Add the arguments to the parser
  ap.add_argument("-i", "--inputFile", required=True,
   help="file from which message to be read")
  ap.add_argument("-se", "--s_email", required=True,
   help="outlook email of sender ")
  ap.add_argument("-p", "--password", required=True,
   help="outlook password")
  ap.add_argument("-re", "--r_email", required=True,
   help="recipient email")
  ap.add_argument("-m", "--message", required=True,
   help="your message for recipient")
  
  args = vars(ap.parse_args())
  # open message
  msg = extract_msg.Message(args['inputFile'])
  # print sender name
  print('Sender: {}'.format(msg.sender))
  # print date
  print('Sent On: {}'.format(msg.date))
  # print subject
  print('Subject: {}'.format(msg.subject))
  # print body
  print('Body: {}'.format(msg.body))
  # save all message and attachments
  msg.save()

  #sending message via outlook
  credentials = Credentials(username=args['s_email'], password=args['password'])
  a = Account(
     primary_smtp_address=args['s_email'], credentials=credentials,
     autodiscover=True, access_type=DELEGATE
     )
  m = Message(
    account=a,
    subject='Daily motivation',
    body=args['message'],
    to_recipients=[
      Mailbox(email_address=args['r_email']),
      ]
      )
  m.send()


if __name__ == "__main__":
    main()

