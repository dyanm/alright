from shutil import make_archive
from loguru import logger
logger.add("Logs/logs.txt", backtrace=True, diagnose=True, level="DEBUG")
logger.debug("Importing libraries")
from alright import WhatsApp
import pandas as pd
from time import sleep
import os
import sys
'''
For Upgrading
cd alright
git pull
python setup.py install --user

For Initial Installation
git clone https://github.com/dyanm/alright.git
cd alright
python setup.py install --user
'''

class VideoTooBigError(Exception):
    pass

@logger.catch(onerror=lambda _: sys.exit(1))
def main():
    logger.debug("Parsing arguments [excel_file, initial_delay(5), task_delay(2)]")
    import argparse
    parser = argparse.ArgumentParser(description='Sends WhatsApp messages in bulk, sending one message per row in the specified excel file')
    parser.add_argument('input_file',type=str, help='The excel file to take the input from.')
    parser.add_argument(
        '-i', '--initial-delay', type=int, default="5", metavar="SECONDS",
        help='The number of seconds to wait when starting up WhatsApp. Set this longer if your WhatsApp takes very long to load/sync. Note that this is an extra delay on top of the WhatsApp automation library\'s own internal delay.'
        )
    parser.add_argument(
        '-d', '--task-delay', type=int, default="2", metavar="SECONDS",
        help='The number of seconds to wait between each major action (finding contact, sending message, sending image). Note that this is an extra delay on top of the WhatsApp automation library\'s own internal delay.'
        )
    args = parser.parse_args()
    
    logger.debug(f"Loading excel file: {args.input_file}")
    df = pd.read_excel(args.input_file)

    logger.debug("Checking video file sizes (WhatsApp has a video size limit of 14MB)")
    files_over_size_limit = []
    for row in df.itertuples(index=False, name=None):
        _, _, _, _, _, video_filename, _, _, _, = row  
        if pd.isnull(video_filename):
            continue
        video_filepath = os.path.join(os.getcwd(), "Videos", video_filename)
        file_size = os.path.getsize(video_filepath)
        file_size_in_mb = file_size / 1024.0 / 1024.0
        if (file_size_in_mb >= 14) and (video_filename not in files_over_size_limit): # No need to warn again if the file is the same
            files_over_size_limit.append(video_filename)
            logger.warning(f"Video file {video_filename} larger than 14MB")
    
    if len(files_over_size_limit) > 0:
        raise VideoTooBigError("At least 1 video file is too big (over 14MB)! Please see the warning messages above.")

    logger.info("Opening WhatsApp with Selenium")
    messenger = WhatsApp()
    logger.info(f"Waiting for WhatsApp to load/sync")
    sleep(args.initial_delay)

    count = 0
    for row in df.itertuples(index=False, name=None):
        count += 1
        logger.debug(f"Starting Task #{count}...")
        name, contact_name, message, img, img_caption, vid, vid_caption, file, file_caption = row
        try: 
            message = message.format(name=name)
        except:
            pass
        
        logger.debug(f"Searching for \"{contact_name}\" in your Contact List")
        isContactFound = messenger.find_by_username(contact_name)
        if not isContactFound:
            logger.debug(f"Unable to find \"{contact_name}\". Skipping to next contact.")
            continue
        else:
            logger.debug(f"\"{contact_name}\" found.")
            
        if not pd.isnull(message):
            logger.debug(f"Sending message to {name}...")
            messenger.send_message(message)
            messenger.wait_until_message_successfully_sent()
            sleep(args.task_delay)

        if not pd.isnull(img):
            image_filepath = os.path.join(os.getcwd(), "Images", img)
            print(image_filepath)
            if not pd.isnull(img_caption):
                logger.debug(f"Sending image with caption to {name}...")
                messenger.send_picture(image_filepath, img_caption)
            else:
                logger.debug(f"Sending image to {name}...")
                messenger.send_picture(image_filepath)
            messenger.wait_until_message_successfully_sent()
            sleep(args.task_delay)

        if not pd.isnull(vid):
            video_filepath = os.path.join(os.getcwd(), "Videos", vid)
            print(video_filepath)
            if not pd.isnull(vid_caption):
                logger.debug(f"Sending video with caption to {name}...")
                messenger.send_video(video_filepath, vid_caption)
            else:
                logger.debug(f"Sending video to {name}...")
                messenger.send_video(video_filepath)
            messenger.wait_until_message_successfully_sent()
            sleep(args.task_delay)
        
        if not pd.isnull(file):
            file_filepath = os.path.join(os.getcwd(), "Files", file)
            print(file_filepath)
            if not pd.isnull(file_caption):
                logger.debug(f"Sending file with caption to {name}...")
                messenger.send_file(file_filepath, file_caption)
            else:
                logger.debug(f"Sending file to {name}...")
                messenger.send_file(file_filepath)
            messenger.wait_until_message_successfully_sent()
            sleep(args.task_delay)

        logger.debug(f"Task #{count} completed. Moving on to next task.")
        sleep(1)
    logger.info(f"All tasks are completed. Automatically exiting in 3 seconds...")
    sleep(3)

if __name__ == "__main__":
    main()
