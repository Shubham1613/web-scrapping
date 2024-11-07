import requests
import sys
import os
import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Set up Selenium WebDriver
def init_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")  # Run in headless mode
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

# Function to fetch all video links from the channel page
def fetch_video_links(driver, channel_url):
    driver.get(channel_url + "/videos")
    time.sleep(1)  # Allow the page to load
    
    # Scroll to load more videos if necessary
    video_links = []
    soup = BeautifulSoup(driver.page_source, "html.parser")
    links = [a["href"] for a in soup.find_all("a", {"id": "video-title-link"})]
    video_links.extend(links)
    driver.execute_script("window.scrollTo(0, document.documentElement.scrollHeight);")

    video_links = list(set(video_links))  # Remove duplicates
    return video_links

# Function to fetch video details
def fetch_video_details(driver, video_url):
    driver.get(video_url)
    time.sleep(1)  # Allow page to load

    # Use BeautifulSoup to parse video details
    soup = BeautifulSoup(driver.page_source, "html.parser")
    title = soup.find("h1", {"class": "ytd-watch-metadata"}).text.strip() if soup.find("h1", {"class": "ytd-watch-metadata"}) else ""
    view_count = soup.find("span", {"class": "bold style-scope yt-formatted-string"}).text.split(' ')[0] if soup.find("span", {"class": "bold style-scope yt-formatted-string"}) else ""
    description = soup.find("span", {"class": "yt-core-attributed-string--link-inherit-color"}).text if soup.find("span", {"class": "yt-core-attributed-string--link-inherit-color"}) else ""
    return {
        "Video ID" : video_url.split('?v=')[1],
        "Video URL": video_url,
        "Title": title,
        "View Count": view_count,
        "Description": description
    }

# Function to fetch comments
def fetch_comments(driver, video_url):
    driver.get(video_url)
    time.sleep(1)  # Allow page to load

    # Scroll to load comments
    driver.execute_script("window.scrollTo(0, document.documentElement.scrollHeight);")
    time.sleep(1)

    # Extract comments using BeautifulSoup
    soup = BeautifulSoup(driver.page_source, "html.parser")
    comments_data = []
    for comment in soup.find_all("div", {"id" : "body"})[:100]:
        author = comment.find("span", {"class": "ytd-comment-view-model"}).text.strip() if comment.find("span", {"class": "ytd-comment-view-model"}) else ""
        comment_text = comment.find("span", {"class" : "yt-core-attributed-string yt-core-attributed-string--white-space-pre-wrap"}).text.strip()
        like_count = comment.find("span", {"id": "vote-count-middle"}).text.strip()
        published_date = comment.find("a", {"class": "yt-simple-endpoint style-scope ytd-comment-view-model"}).text.strip()

        comments_data.append({
            "Video URL": video_url,
            "Comment Text": comment_text,
            "Author": author,
            "Published Date": published_date,
            "Like Count": like_count,
        })

    return comments_data

# Save data to Excel
def save_to_excel(videos_data, comments_data, filename="YouTube_Data.xlsx"):
    with pd.ExcelWriter(filename) as writer:
        # Video data
        pd.DataFrame(videos_data).to_excel(writer, sheet_name="Video Data", index=False)
        
        # Comments data
        pd.DataFrame(comments_data).to_excel(writer, sheet_name="Comments Data", index=False)


def main(channel_url):
    driver = init_driver()
    try:
        print("Fetching video links...")
        video_links = fetch_video_links(driver, channel_url)
        
        # Fetch video details
        videos_data = []
        comments_data = []
        for video_link in video_links:
            video_url = "https://www.youtube.com" + video_link
            print(f"Fetching data for video: {video_url}")
            video_data = fetch_video_details(driver, video_url)
            videos_data.append(video_data)

            # Fetch comments for each video
            video_comments = fetch_comments(driver, video_url)
            comments_data.extend(video_comments)

        # Save to Excel
        print("Saving data to Excel...")
        save_to_excel(videos_data, comments_data)
        print("Data saved successfully.")
    finally:
        driver.quit()


if __name__ == '__main__':

    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--url", type=str, help="URL of youtube channel")
    args = parser.parse_args()

    if args.url:
        main(args.url)

