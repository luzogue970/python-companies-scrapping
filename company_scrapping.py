from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd

driver = webdriver.Chrome()

activities = ["55.10Z", "56.21Z"]
departments = ["35", "44", "22", "56", "29", ]
min_people = "6"

dictionary = {}


def find_companies():
    companies = driver.find_elements(By.CLASS_NAME, "container-resultat")
    if len(companies) < 1:
        return
    for company in companies:
        try:
            company.find_element(By.CLASS_NAME, "entreprise-cessee")
            continue
        except:
            company_name = company.find_element(By.CLASS_NAME, "nom-entreprise").text
            if company_name in dictionary:
                continue
            my_activity = \
                company.find_element(By.CLASS_NAME, "fiche-entreprise").find_elements(By.CLASS_NAME, "content")[
                    1].find_element(By.CLASS_NAME, "value").text
            number_of_people = \
            company.find_element(By.CLASS_NAME, "fiche-entreprise").find_elements(By.CLASS_NAME, "content")[
                3].find_elements(By.CLASS_NAME, "key")[0].find_element(By.TAG_NAME, "span").text
            dictionary[company_name] = {
                "company_name": company_name,
                "domaine": my_activity,
                "effectif": number_of_people,
                "departement": department
            }


for activity in activities:
    for department in departments:
        driver.get(
            f"https://www.pappers.fr/recherche?ciblerActivitePrincipale=true&departement={department}&activite={activity}&effectifs_min={min_people}"
        )
        time.sleep(2)
        number_of_pages_string = driver.find_element(By.CLASS_NAME, "texte-droite" and "desktop-only").text.split(" ")[
            -1]
        try:
            test_int = int(number_of_pages_string)
            for i in range(int(number_of_pages_string)):
                time.sleep(1)
                find_companies()
                if int(number_of_pages_string) == i + 1:
                    break
                time.sleep(2)
                page_change = driver.find_element(By.CLASS_NAME, "pagination" and "pagination-image-right")
                page_change.click()
        except:
            find_companies()

df_existing = pd.DataFrame(columns=['Nom Entreprise', "Domaine d'activité", "effectif", "departement"])
data_list = []
for company, details in dictionary.items():
    data_list.append([details['company_name'], details['domaine'], details["effectif"], details['departement']])
df_new = pd.DataFrame(data_list, columns=['Nom Entreprise', "Domaine d'activité", "effectif", 'departement'])
df_combined = pd.concat([df_existing, df_new], ignore_index=True)
df_combined.to_excel("entreprises_alternance.xlsx", index=False, sheet_name="entreprises_alternance")
