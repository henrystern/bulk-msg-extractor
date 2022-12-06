import os
import extract_msg
import pandas as pd

def extract_message_contents(file, directory="."):
    # returns dictionary of key data from a message at directory/file
    path = f"{directory}/{file}"
    msg = extract_msg.Message(path)
    msg = msg.getJson()
    msg = pd.read_json(msg, typ='series')
    # long formulas cant be written to xlsx so if too long print path as string
    # TODO this should create the link relative to output path not the script working directory
    # set to 1 because it was faster to just fix the long links in vba from strings
    if len(path) < 1: 
        contents = {"file": f"=hyperlink(\"{path}\")"}
    else:
        contents = {"file": path}
    contents.update(msg)
    return contents

def localize_naive_datetime(series, tz):
    # makes datetime tz unaware with localized time
    return series.dt.tz_convert(tz).dt.tz_localize(None)

def recursive_extract_emails(directory, output_directory="."):
    # recurse through directory and append each emails data to all_email list
    all_emails = []
    for subdir, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith(".msg"):
                file_contents = extract_message_contents(file, subdir)
                all_emails.append(file_contents)
    return all_emails


directory = '' # path to root directory of messages

output_directory = 'out'
output_name = 'All_Emails.xlsx'

raw_data = recursive_extract_emails(directory)

# create df and set dtypes for easier handling in excel
df = pd.DataFrame(raw_data)
df = df.set_index('file')
df['date'] = pd.to_datetime(df['date'], utc=True)
df['date'] = localize_naive_datetime(df['date'], "US/Eastern")

# drop duplicates
df = df.drop_duplicates(subset=["from", "date", "body"])

# write df to xlsx
with pd.ExcelWriter(
    f"{output_directory}/{output_name}",
    datetime_format="YYYY-MM-DD HH:MM:SS"
) as writer:
    df.to_excel(writer)
