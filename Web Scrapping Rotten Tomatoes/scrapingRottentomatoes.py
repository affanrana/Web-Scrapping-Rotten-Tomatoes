from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "Most Popular TV Shows This Month"
sheet.append(["Rank", "TV Show Name", "Critic Score", "Audience Score"])

try:
    sourcey = requests.get("https://www.rottentomatoes.com/browse/tv_series_browse/sort:popular")
    sourcey.raise_for_status()

    soupy = BeautifulSoup(sourcey.text, "html.parser")

    shows = soupy.find('div', class_="discovery-tiles__wrap").find_all("div", class_="flex-container")

    ranker = 1
    for show in shows:
        name = show.find("span", class_="p--small").text.strip()
        showInfo = show.find('score-pairs-deprecated')
        audienceScore = showInfo['audiencescore']
        criticScore = showInfo['criticsscore']
        print(str(ranker) + ". " + name + " - Critic Score: " + criticScore + " - Audience Score: " + audienceScore)
        sheet.append([str(ranker), name, criticScore, audienceScore])
        ranker += 1
except Exception as e:
    print(e)

excel.save("Rottentomatoes Most Popular TV Shows This Month.xlsx")