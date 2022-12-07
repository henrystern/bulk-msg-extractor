import os
import extract_msg
import pandas as pd


def main(search_directory, output_path):
    print("Processing...")
    raw_data = recursive_extract_emails(search_directory)

    # create df and set dtypes for easier handling in excel
    df = pd.DataFrame(raw_data)
    df = df.set_index('file')
    df['date'] = pd.to_datetime(df['date'], utc=True)
    df['date'] = localize_naive_datetime(df['date'], "US/Eastern")

    # drop duplicates
    df = df.drop_duplicates(subset=["from", "date", "body"])
    print("\nWriting to file...")
    # write df to xlsx
    with pd.ExcelWriter(
        f"{output_path}",
        datetime_format="YYYY-MM-DD HH:MM:SS"
    ) as writer:
        df.to_excel(writer)

    print("Done!")


def extract_message_contents(file, search_directory="."):
    # returns dictionary of key data from a message at search_directory/file
    path = f"{search_directory}\\{file}"
    msg = extract_msg.Message(path)
    msg = msg.getJson()
    msg = pd.read_json(msg, typ='series')
    # long formulas cant be written to xlsx so if too long print path as string
    # TODO this should create the link relative to output path not the script working directory
    file_link = f"(path too long) {path}" if len(
        path) > 210 else f"=hyperlink(\"{path}\", \"link\")"
    directory_link = f"(path too long) {search_directory}" if len(
        path) > 210 else f"=hyperlink(\"{search_directory}\", \"link\")"
    contents = {"file": file, "file_link": file_link,
                "directory": search_directory, "directory_link": directory_link}
    contents.update(msg)
    return contents


def localize_naive_datetime(series, tz):
    # makes datetime tz unaware with localized time
    return series.dt.tz_convert(tz).dt.tz_localize(None)


def recursive_extract_emails(search_directory):
    # recurse through directory and append each emails data to all_email list
    all_emails = []
    old_subdir = ""
    print("Scanning: ", end="\r")
    for subdir, _, files in os.walk(search_directory):
        #progress TODO doesn't work for some strings, seems to be length related
        print(" " * (len(old_subdir) + 10), end="\r")
        print(f"Scanning: {subdir}", end="\r")
        old_subdir = subdir
        
        for file in files:
            if file.endswith(".msg"):
                file_contents = extract_message_contents(file, subdir)
                all_emails.append(file_contents)
    return all_emails


if __name__ == "__main__":
    output_path = input(
        "Enter the path of the output file (absolute or relative to script path, default .\\out\\extracted_msg.xlsx): ") or ".\\out\\extracted_msg.xlsx"
    os.chdir(output_path)
    search_directory = input(
        "Enter the path of the search directory (absolute or relative to output path, default .\\Data): ") or ".\\Data"
    main(search_directory, output_path)