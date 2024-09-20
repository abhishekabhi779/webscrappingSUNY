import requests
from bs4 import BeautifulSoup
import pandas as pd
from io import StringIO
import xlsxwriter

global_id_counter = 1


BASE_URL = "https://banner.sunypoly.edu"
START_PATH = "/pls/prod/swssschd.P_SelDefSchTerm"

def fetch_terms(session):
    """Fetch and return available terms from the dropdown."""
    response = session.get(f"{BASE_URL}{START_PATH}")
    if response.status_code != 200:
        print("Failed to fetch terms: Status Code", response.status_code)
        return []
    soup = BeautifulSoup(response.text, 'html.parser')
    return [(term.text.strip(), term['value']) for term in soup.select("select[name='term_in'] option")]

def fetch_disciplines(session, term_value):
    """Fetch and return disciplines for a given term."""
    response = session.post(f"{BASE_URL}/pls/prod/swssschd.P_SelDisc", data={'term_in': term_value})
    if response.status_code != 200:
        print("Failed to fetch disciplines: Status Code", response.status_code)
        return []
    soup = BeautifulSoup(response.text, 'html.parser')
    return [(disc.text.strip(), disc['value']) for disc in soup.select("select[name='disc_in'] option")]


def fetch_course_schedule(session, term_value, discipline_value):
    global global_id_counter
    """Fetch and return schedule data as a DataFrame."""
    form_data = {
        'term_in': term_value,
        'disc_in': discipline_value
    }
    action_path = "/pls/prod/swssschd.P_ShowSchd"  # Correct path based on form action
    response = session.post(f"{BASE_URL}{action_path}", data=form_data)
    
    print("Scraping data for term:", term_value, "and discipline:", discipline_value)
    print("Response Status Code:", response.status_code)  # Check if the request was successful

    if response.status_code == 200:
        try:
            tables = pd.read_html(StringIO(response.text), header=0)
            if tables:
                df = tables[2]  # Assume the first table is the relevant one
                
                # Add Unique ID column
                num_rows = len(df)
                df['ID'] = range(global_id_counter, global_id_counter + num_rows)
                global_id_counter += num_rows
                
                # Add Term column
                df['Term'] = term_value
                
                return df
            else:
                print("No tables found in the HTML.")
                return pd.DataFrame()
        except Exception as e:
            print("An error occurred while parsing tables:", e)
            return pd.DataFrame()
    else:
        print("Failed to fetch course schedule data. Status Code:", response.status_code)
        return pd.DataFrame()


def save_to_excel(data_frames, filename="output3.xlsx"):
    """Save data frames to an Excel file, each frame in a separate sheet."""
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        for term_value, df in data_frames.items():
            sheet_name = term_value[:31]  # Ensure sheet name does not exceed Excel's limit
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print("Excel file created:", filename)


def main():
    session = requests.Session()
    terms = fetch_terms(session)
    all_data = {}

    if not terms:
        print("No terms available to process.")
        return

    for term_name, term_value in terms:
        print(f"Processing term: {term_name}({term_value})")
        disciplines = fetch_disciplines(session, term_value)
        term_data = []

        for disc_name, disc_value in disciplines:
            print(f"Scraping schedule for {term_name} - {disc_name}")
            df = fetch_course_schedule(session, term_value, disc_value)
            if not df.empty:
                
                term_data.append(df)

        if term_data:
            all_data[term_value] = pd.concat(term_data, ignore_index=True)

    if all_data:
        save_to_excel(all_data)
    else:
        print("No data collected to write to Excel.")

if __name__ == "__main__":
    main()
