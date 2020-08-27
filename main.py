import extract_msg
import sys
import argparse


sys.argv = ["test.py", "-i", "C:\\Users\\msoln\\OneDrive\\Documents\\test.msg", "-t", "msg"]


def parse_outlook_email(email_file_path, email_file_type):
    email_file = email_file_path
    msg = extract_msg.Message(email_file)
    msg_sender = msg.sender
    msg_message = msg.body
    msg_header = msg.headerDict

    print('Header: {}'.format(msg_header))
    print('Sender: {}'.format(msg_sender))
    print('Body: {}'.format(msg_message))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', action='store', dest='email_file_path',
                        help='path of the email file', required=True)

    parser.add_argument('-t', action='store', dest='email_file_type',
                        help='email file type', choices=['msg'], required=True)

    args = parser.parse_args()

    if args.email_file_type == "msg":
        parse_outlook_email(args.email_file_path, args.email_file_type)


if __name__ == "__main__":
    main()
