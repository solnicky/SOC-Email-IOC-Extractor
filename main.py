import extract_msg
import sys
import argparse
import re


sys.argv = ["test.py", "-i", "C:\\Users\\msoln\\OneDrive\\Documents\\test.msg", "-t", "msg"]


def extract_urls(msg_body):
    regex = r"https?://[^\s]+"
    urls = re.findall(regex, msg_body)
    return [url for url in urls]


def extract_true_sender(msg_header):
    return "\n".join(msg_header['Received'])


def parse_outlook_email(email_file_path):
    email_file = email_file_path
    msg = extract_msg.Message(email_file)
    msg_sender = msg.sender
    msg_body = msg.body
    msg_header = msg.headerDict

    true_sender = extract_true_sender(msg_header)
    urls = "\n".join(extract_urls(msg_body))

    return [true_sender, urls]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', action='store', dest='email_file_path',
                        help='path of the email file', required=True)

    parser.add_argument('-t', action='store', dest='email_file_type',
                        help='email file type', choices=['msg'], required=True)

    args = parser.parse_args()

    if args.email_file_type == "msg":
        return parse_outlook_email(args.email_file_path)


if __name__ == "__main__":
    for x in main():
        print(x)
