"""
This script was inspired from tmcw's Ruby script doing the same thing:

    https://gist.github.com/tmcw/1098861

And recent fixes implemented thanks to the login structure by wederbrand:

    https://github.com/wederbrand/workout-exchange/blob/master/garmin_connect/download_all.rb

The goal is to iteratively download all detailed information from Garmin Connect
and store it locally for further perusal and analysis. This is still very much
preliminary; future versions should include the ability to seamlessly merge
all the data into a single file, filter by workout type, and other features
to be determined.

2018-04-11 - Garmin appears to have deprecated its old REST api and legacy authentication
The following updates work for me using Python 2.7 and Mechanize
"""

import argparse
from getpass import getpass
import json
import mechanize as me
import os
import re
import shutil
import sys
import urllib
import datetime
import string
import zipfile
import subprocess

BASE_URL = "https://sso.garmin.com/sso/login"
GAUTH = "https://connect.garmin.com/modern/auth/hostname"
SSO = "https://sso.garmin.com/sso"
CSS = "https://static.garmincdn.com/com.garmin.connect/ui/css/gauth-custom-v1.2-min.css"
REDIRECT = "https://connect.garmin.com/modern/"
ACTIVITIES = "https://connect.garmin.com/modern/proxy/activitylist-service/activities/search/activities?start=%s&limit=%s"
WELLNESS = "https://connect.garmin.com/modern/proxy/userstats-service/wellness/daily/%s?fromDate=%s&untilDate=%s"
DAILYSUMMARY = "https://connect.garmin.com/modern/proxy/wellness-service/wellness/dailySummaryChart/%s?date=%s"

ORIGINAL = "https://connect.garmin.com/modern/proxy/download-service/files/activity/%s"
TCX = "https://connect.garmin.com/modern/proxy/download-service/export/tcx/activity/%s"
GPX = "https://connect.garmin.com/modern/proxy/download-service/export/gpx/activity/%s"

def login(agent, username, password):
    global BASE_URL, GAUTH, REDIRECT, SSO, CSS

    # First establish contact with Garmin and decipher the local host.
    agent.set_handle_robots(False)   # no robots
    page = agent.open(BASE_URL)
    pattern = "\'\S+sso\.garmin\.com\\S+\'"
    script_url = re.search(pattern, page.get_data()).group()[1:-1]
    agent.set_handle_robots(False)   # no robots
    agent.set_handle_refresh(False)  # can sometimes hang without this
    agent.open(script_url)
    agent.addheaders = [('User-agent', 'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.181 Safari/537.36')]
    hostname_url = agent.open(GAUTH)
    hostname = json.loads(hostname_url.get_data())['host']

    # Package the full login GET request...
    data = {'service': REDIRECT,
        'webhost': hostname,
        'source': BASE_URL,
        'redirectAfterAccountLoginUrl': REDIRECT,
        'redirectAfterAccountCreationUrl': REDIRECT,
        'gauthHost': SSO,
        'locale': 'en_US',
        'id': 'gauth-widget',
        'cssUrl': CSS,
        'clientId': 'GarminConnect',
        'rememberMeShown': 'true',
        'rememberMeChecked': 'false',
        'createAccountShown': 'true',
        'openCreateAccount': 'false',
        'usernameShown': 'false',
        'displayNameShown': 'false',
        'consumeServiceTicket': 'false',
        'initialFocus': 'true',
        'embedWidget': 'false',
        'generateExtraServiceTicket': 'false'}

    # ...and officially say "hello" to Garmin Connect.
    login_url = 'https://sso.garmin.com/sso/login?%s' % urllib.urlencode(data)
    agent.open(login_url)

    # Set up the login form.
    agent.select_form(predicate = lambda f: 'id' in f.attrs and f.attrs['id'] == 'login-form')
    agent['username'] = username
    agent['password'] = password
    agent.set_handle_robots(False)   # no robots
    agent.set_handle_refresh(False)  # can sometimes hang without this
    # Apparently Garmin Connect attempts to filter on these browser headers;
    # without them, the login will fail.

    # Submit the login!
    res = agent.submit()
    if res.get_data().find("Invalid") >= 0:
        quit("Login failed! Check your credentials, or submit a bug report.")
    elif res.get_data().find("SUCCESS") >= 0:
        print('Login successful! Proceeding...')
    else:
        quit('UNKNOWN STATE. This script may need to be updated. Submit a bug report.')

    # Now we need a very specific URL from the response.
    response_url = re.search("response_url\s*=\s*\"(.*)\";", res.get_data()).groups()[0]
    agent.open(response_url.replace("\/", "/"))

    # In theory, we're in.

def file_exists_in_folder(filename, folder):
    "Check if the file exists in folder of any subfolder"
    for _, _, files in os.walk(folder):
        if filename in files:
            return True
    return False

def garmin_encode_file_name(dt):
    table = string.digits + string.ascii_uppercase
    epoch = datetime.datetime(year=1989, month=12, day=31)  # timezone omitted

    name = str(dt.year - 2010)
    name += table[dt.month]
    name += table[dt.day]
    name += table[dt.hour]
    name += str(int((dt - epoch).total_seconds()))[-4:]

    return name

def activities(agent, outdir, increment = 100):
    global ACTIVITIES
    currentIndex = 0
    initUrl = ACTIVITIES % (currentIndex, increment)  # 100 activities seems a nice round number
    try:
        response = agent.open(initUrl)
    except:
        print('Wrong credentials for user {}. Skipping.'.format(username))
        return
    search = json.loads(response.get_data())
    while True:
        if len(search) == 0:
            # All done!
            print('Download complete')
            break

        for item in search:
            # Read this list of activities and save the files.

            activityId = item['activityId']
            activityDate = item['startTimeLocal'][:10] # startTimeLocal: "YYYY-MM-DD hh:mm:ss"
            url = ORIGINAL % activityId
            file_name = '{}_{}.zip'.format(activityDate, activityId)
            if file_exists_in_folder(file_name, output):
                print('{} already exists in {}. Skipping.'.format(file_name, output))
                continue
            print('{} is downloading...'.format(file_name))
            datafile = agent.open(url).get_data()
            file_path = os.path.join(outdir, file_name)
            f = open(file_path, "wb")
            f.write(datafile)
            f.close()
            #shutil.copy(file_path, os.path.join(os.path.dirname(os.path.dirname(file_path)), file_name))

            z = zipfile.ZipFile(file_path, 'r')
            if len(z.infolist()) != 1:
                raise ValueError('zip file expected to contain a single entry: ' + file_path)
            zinfo = z.infolist()[0]
            z.extract(zinfo, os.path.dirname(file_path))
            z.close()
            del z
            os.remove(file_path)

            if len(item['startTimeLocal']) != 19:
                raise ValueError('unsupported activity date: ' + item['startTimeLocal'])
            activityDateTime = datetime.datetime(
                year=int(item['startTimeLocal'][0:4]),
                month=int(item['startTimeLocal'][5:7]),
                day=int(item['startTimeLocal'][8:10]),
                hour=int(item['startTimeLocal'][11:13]),
                minute=int(item['startTimeLocal'][14:16]),
                second=int(item['startTimeLocal'][17:19]))
            garmin_encoded_fit_name = unicode(garmin_encode_file_name(activityDateTime))
            final_file_name = '{:04}{:02}{:02}_{:02}{:02}{:02}_{}_{}{}'.format(
                activityDateTime.year,
                activityDateTime.month,
                activityDateTime.day,
                activityDateTime.hour,
                activityDateTime.minute,
                activityDateTime.second,
                garmin_encoded_fit_name,
                activityId,
                os.path.splitext(zinfo.filename)[1].upper())
            final_file = os.path.join(os.path.dirname(file_path), final_file_name)

            shutil.move(
                os.path.join(os.path.dirname(file_path), zinfo.filename),
                final_file)

            # change file times to activity date
            nircmd_date = '{:02}-{:02}-{:04} {:02}:{:02}:{:02}'.format(
                activityDateTime.day,
                activityDateTime.month,
                activityDateTime.year,
                activityDateTime.hour,
                activityDateTime.minute,
                activityDateTime.second)
            subprocess.call(
                [r'D:\prog\apps\nircmd64\nircmdc.exe',
                    'setfiletime',
                    final_file,
                    nircmd_date, nircmd_date, nircmd_date],
                env={}, shell=False)

        # We still have at least 1 activity.
        currentIndex += increment
        url = ACTIVITIES % (currentIndex, increment)
        response = agent.open(url)
        search = json.loads(response.get_data())

def wellness(agent, start_date, end_date, display_name, outdir):
    url = WELLNESS % (display_name, start_date, end_date)
    try:
        response = agent.open(url)
    except:
        print('Wrong credentials for user {}. Skipping.'.format(username))
        return
    content = response.get_data()

    file_name = '{}_{}.json'.format(start_date, end_date)
    file_path = os.path.join(outdir, file_name)
    with open(file_path, "w") as f:
        f.write(content)


def dailysummary(agent, date, display_name, outdir):
    url = DAILYSUMMARY % (display_name, date)
    try:
        response = agent.open(url)
    except:
        print('Wrong credentials for user {}. Skipping.'.format(username))
        return
    content = response.get_data()

    file_name = '{}_summary.json'.format(date)
    file_path = os.path.join(outdir, file_name)
    with open(file_path, "w") as f:
        f.write(content)


def login_user(username, password):
    # Create the agent and log in.
    agent = me.Browser()
    print("Attempting to login to Garmin Connect...")
    login(agent, username, password)
    return agent


def download_files_for_user(agent, username, output):
    user_output = os.path.join(output, username)
    download_folder = os.path.join(user_output, 'Historical')

    # Create output directory (if it does not already exist).
    if not os.path.exists(download_folder):
        os.makedirs(download_folder)

    # Scrape all the activities.
    activities(agent, download_folder)


def download_wellness_for_user(agent, username, start_date, end_date, display_name, output):
    user_output = os.path.join(output, username)
    download_folder = os.path.join(user_output, 'Wellness')

    # Create output directory (if it does not already exist).
    if not os.path.exists(download_folder):
        os.makedirs(download_folder)

    # Scrape all wellness data.
    wellness(agent, start_date, end_date, display_name, download_folder)
    # Daily summary does not do ranges, only fetch for `startdate`
    dailysummary(agent, start_date, display_name, download_folder)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description = 'Garmin Data Scraper',
        epilog = 'Because the hell with APIs!', add_help = 'How to use',
        prog = 'python download.py [-u <user> | -c <csv fife with credentials>] [ -s <start_date> -e <end_date> -d <display_name> ] -o <output dir>')
    parser.add_argument('-u', '--user', required = False,
        help = 'Garmin username. This will NOT be saved!',
        default = None)
    parser.add_argument('-c', '--csv', required = False,
        help = 'CSV file with username and password in "username,password" format.',
        default = None)
    parser.add_argument('-s', '--startdate', required = False,
        help = 'Start date for wellness data (YYYY-MM-DD)',
        default = None)
    parser.add_argument('-e', '--enddate', required = False,
        help = 'End date for wellness data (YYYY-MM-DD)',
        default = None)
    parser.add_argument('-d', '--displayname', required = False,
        help = 'Displayname (see the url when logged into Garmin Connect)',
        default = None)
    parser.add_argument('-o', '--output', required = False,
        help = 'Output directory.', default = os.path.join(os.getcwd(), 'Results/'))
    args = vars(parser.parse_args())

    # Sanity check, before we do anything:
    if args['user'] is None and args['csv'] is None:
        print("Must either specify a username (-u) or a CSV credentials file (-c).")
        sys.exit()

    # Try to use the user argument from command line
    output = args['output']

    if args['user'] is not None:
        password = getpass('Garmin account password (NOT saved): ')
        username = args['user']
    else:
        csv_file_path = args['csv']
        if not os.path.exists(csv_file_path):
            print("Could not find specified credentials file \"{}\"".format(csv_file_path))
            sys.exit()
        try:
            with open(csv_file_path, 'r') as f:
                contents = f.read()
        except IOError as e:
            print(e)
            sys.exit()
        try:
            username, password = contents.strip().split(",")
        except IndexError:
            print("CSV file must only have 1 line, in \"username,password\" format.")
            sys.exit()

    # Perform the download.
    if args['startdate'] is not None:
        start_date = args['startdate']
        end_date = args['enddate']
        display_name = args['displayname']
        if not end_date:
            print("Provide an enddate")
            sys.exit(1)
        if not display_name:
            print("Provide a displayname, you can find it in the url of Daily Summary: '.../daily-summary/<displayname>/...'")
            sys.exit(1)
        agent = login_user(username, password)
        download_wellness_for_user(agent, username, start_date, end_date, display_name, output)
    else:
        agent = login_user(username, password)
        download_files_for_user(agent, username, output)
