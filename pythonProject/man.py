import requests
from googleapiclient.discovery import build
from bs4 import BeautifulSoup


def google_search(query):
    # Your actual API Key and Custom Search Engine ID
    api_key = 'AIzaSyAo8QiTOZAgQnCa9eaFh9E2bZVsfnj8UQA'  # Replace with your API Key
    cse_id = 'a203374f897c743d4'  # Replace with your Custom Search Engine ID

    # Build the service for Custom Search API
    service = build("customsearch", "v1", developerKey=api_key)

    try:
        # Perform the search
        res = service.cse().list(q=query, cx=cse_id).execute()

        # Parse and print the results
        if 'items' in res:
            for item in res['items']:
                title = item['title']
                link = item['link']
                print(f"Title: {title}")
                print(f"Link: {link}")


                # Fetch and parse the page's inner content
                page_content = fetch_page_content(link)
                print(f"Inner content: {page_content[:300]}...\n")  # Print first 300 characters of content
        else:
            print("No results found.")
    except Exception as e:
        print(f"An error occurred: {e}")


def fetch_page_content(url):
    try:
        # Send an HTTP request to fetch the page content
        response = requests.get(url)

        # If the request is successful (status code 200)
        if response.status_code == 200:
            # Parse the page content with BeautifulSoup
            soup = BeautifulSoup(response.text, 'html.parser')

            # Extract the main text content (you can refine this based on the structure of the page)
            # For example, we are extracting all paragraphs:
            paragraphs = soup.find_all('p')
            text_content = ' '.join([para.get_text() for para in paragraphs])

            return text_content
        else:
            return "Failed to retrieve the page"
    except Exception as e:
        return f"An error occurred while fetching the page: {e}"


# Query to search for
query = "which Article of Indian Constitution provides establishment of Panchayat Raj System in India"
google_search(query)
