import os
import extract_msg
import pandas as pd

directory = 'Data'

def extract_message_contents(file, directory="."):
    # returns dictionary of key data from a message at directory/file
    msg = extract_msg.Message(f"{directory}/{file}")
    msg = msg.getJson()
    msg = pd.read_json(msg, typ='series')
    contents = {"file": f"=hyperlink(\"{directory}/{file}\")"}
    contents.update(msg)
    return contents

def localize_naive_datetime(series, tz):
    # makes datetime tz unaware with localized time
    return series.dt.tz_convert(tz).dt.tz_localize(None)

all_emails = []
# recurse through directory and append each emails data to all_email list
for subdir, dirs, files in os.walk(directory):
    for file in files:
        if file.endswith(".msg"):
            file_contents = extract_message_contents(file, subdir)
            all_emails.append(file_contents)

# create df and set dtypes for easier handling in excel
df = pd.DataFrame(all_emails)
df = df.set_index('file')
df['date'] = localize_naive_datetime(pd.to_datetime(df['date'], utc=True), "US/Eastern") 

# write df to xlsx
with pd.ExcelWriter(
    "data.xlsx",
    datetime_format="YYYY-MM-DD HH:MM:SS"
) as writer:
    df.to_excel(writer)  