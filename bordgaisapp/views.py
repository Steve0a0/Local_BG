import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from rest_framework.views import APIView
from rest_framework.parsers import MultiPartParser, FormParser
from rest_framework.response import Response
from rest_framework import status
from django.conf import settings
from .serializers import UploadedFileSerializer
import time

plan_xpath_mapping = {
    'Dual Fuel and €100 Welcome Bonus': "/html/body/div/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[1]/div[2]/div[4]/a",
    'Dual Fuel and Free Hive': "/html/body/div/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[2]/div[2]/div[4]/a",
    'Green Dual Fuel and €100 Welcome Bonus': "/html/body/div/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[3]/div[2]/div[4]/a",
    'EV Smart Dual Fuel and €150 Credit with EV Charger Install': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[1]/div[2]/div[4]/a",
    'Free Saturday Dual Fuel and Free Hive': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[2]/div[2]/div[4]/a",
    'Free Sunday Dual Fuel and Free Hive': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[3]/div[2]/div[4]/a",
    'Free Time Sat Dual Fuel and €100 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[4]/div[2]/div[4]/a",
    'Free Time Sun Dual Fuel and €100 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[5]/div[2]/div[4]/a",
    'Green EV Smart Dual Fuel and €150 Credit with EV Charger Install': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[6]/div[2]/div[4]/a",
    'Green Free Saturday Dual Fuel and €100 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[7]/div[2]/div[4]/a",
    'Green Free Sunday Dual Fuel and €100 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[8]/div[2]/div[4]/a",
    'Green Standard Smart Dual Fuel and €100 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[9]/div[2]/div[4]/a",
    'Green Weekend Dual Fuel and €100 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[10]/div[2]/div[4]/a",
    'Smart EV Dual Fuel and Free Hive': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[11]/div[2]/div[4]/a",
    'Standard Smart Dual Fuel and €100 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[12]/div[2]/div[4]/a",
    'Standard Smart Dual Fuel and Free Hive': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[13]/div[2]/div[4]/a",
    'Standard Smart Dual Fuel Free Hive+Amazon Echo Dot': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[14]/div[2]/div[4]/a",
    'Weekend Dual Fuel and Free Hive': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[15]/div[2]/div[4]/a",
    'Weekend Smart Duel Fuel and €100 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[16]/div[2]/div[4]/a",
    'Electricity Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[1]/div[2]/div[4]/a",
    'Electricity Discount + €50 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[2]/div[2]/div[4]/a",
    'Green Electricity Discount + €50 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[3]/div[2]/div[4]/a",
    'EV New Elec Only and €150 Credit with EV Charger Install': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[1]/div[2]/div[4]/a",
    'EV Smart Electricity Discount + €50 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[2]/div[2]/div[4]/a",
    'Free Time Sat Electricity and €50 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[3]/div[2]/div[4]/a",
    'Free Time Sun Electricity and €50 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[4]/div[2]/div[4]/a",
    'Green EV Smart Electricity and €50 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[5]/div[2]/div[4]/a",
    'Green Free Saturday Electricity and €50 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[6]/div[2]/div[4]/a",
    'Green Free Sunday Electricity and €50 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[7]/div[2]/div[4]/a",
    'Green Smart EV and €150 Credit with EV Charger Install': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[8]/div[2]/div[4]/a",
    'Green Standard Smart Electricity and €50 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[9]/div[2]/div[4]/a",
    'Green Standard Smart Electricity Discount + €50 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[10]/div[2]/div[4]/a",
    'GreenWeekend Smart Electricity Discount + €50 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[11]/div[2]/div[4]/a",
    'Mighty Weekender Smart Electricity Plan': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[12]/div[2]/div[4]/a",
    'Standard Smart Electricity Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[13]/div[2]/div[4]/a",
    'Standard Smart Electricity Discount + €50 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[14]/div[2]/div[4]/a",
    'Standard Smart Electricity Free Hive + Free Amazon Echo Dot': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[15]/div[2]/div[4]/a",
    'Weekend Smart Electricity Discount + €50 Welcome Bonus': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[16]/div[2]/div[4]/a",
    'New Gas Discount and €50 Welcome Credit': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div/div[2]/div[4]/a",
    'Dual Fuel Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[2]/div[2]/div[4]/a",
    'Green Dual Fuel Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[3]/div[2]/div[4]/a",
    'EV Smart Dual Fuel Discount and €150 Credit with EV Charger Install': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[1]/div[2]/div[4]/a",
    'Free Time Sat Dual Fuel Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[4]/div[2]/div[4]/a",
    'Free Time Sun Dual Fuel Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[5]/div[2]/div[4]/a",
    'Green EV Smart Dual Fuel Discount and €150 Credit with EV Charger Install': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[6]/div[2]/div[4]/a",
    'Green Free Saturday Dual Fuel Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[7]/div[2]/div[4]/a",
    'Green Free Sunday Dual Fuel Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[8]/div[2]/div[4]/a",
    'Green Standard Smart Dual Fuel Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[9]/div[2]/div[4]/a",
    'Green Weekend Dual Fuel Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[10]/div[2]/div[4]/a",
    'Standard Smart Duel Fuel Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[13]/div[2]/div[4]/a",
    'Weekend Smart Duel Fuel Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[15]/div[2]/div[4]/a",
    'Green Electricity Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[2]/div[2]/div[4]/a",
    'EV Smart Electricity Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[1]/div[2]/div[4]/a",
    'Free Time Sat Electricty Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[2]/div[2]/div[4]/a",
    'Free Time Sun Electricty Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[3]/div[2]/div[4]/a",
    'Green EV Smart Electricity Discount and €150 Credit with EV Charger Install': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[4]/div[2]/div[4]/a",
    'Green Free Saturday Electricity Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[5]/div[2]/div[4]/a",
    'Green Free Sunday Electricity Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[5]/div[2]/div[4]/a",
    'Green Standard Smart Electricity Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[7]/div[2]/div[4]/a",
    'Green Weekend Smart Electricity Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[8]/div[2]/div[4]/a",
    'Weekend Smart Electricity Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[11]/div[2]/div[4]/a",
    'Weekend Smart Duel Fuel Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[27]/div[2]/div[4]/a",
    'Gas Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[1]/div[2]/div[4]/a",
    'Standard Smart Electricity with Free Hive + Free Amazon Echo Dot': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[15]/div[2]/div[4]/a",
    'Weekend Smart Duel Fuel Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[27]/div[2]/div[4]/a",
    'Standard Smart Dual Fuel Discount': "/html/body/div[1]/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[1]/div[24]/div[2]/div[4]/a"
}


class FileUploadView(APIView):
    parser_classes = (MultiPartParser, FormParser)

    def post(self, request, *args, **kwargs):
        file_serializer = UploadedFileSerializer(data=request.data)
        if file_serializer.is_valid():
            file_instance = file_serializer.save()
            output_path = process_file(file_instance.file)
            return Response({
                "message": "File uploaded and processed successfully!",
                "output_file": output_path
            }, status=status.HTTP_201_CREATED)
        else:
            return Response(file_serializer.errors, status=status.HTTP_400_BAD_REQUEST)


def process_file(uploaded_file):
    df = pd.read_excel(uploaded_file, dtype={'GPRN': str, 'Mobile Number': str, 'joint_account_holder_ Mobile Number':str})


    success_data = []
    failed_data = []

    def initialize_driver():
        driver = webdriver.Chrome()
        driver.get("https://www.bordgaisenergy.ie")
        driver.maximize_window()
        time.sleep(2)
        return driver

    def scroll_into_view(driver, element):
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        time.sleep(0.5)

    def accept_cookies(driver):
        try:
            accept_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div[1]/div/div[2]/div/button[2]"))
            )
            scroll_into_view(driver, accept_button)
            accept_button.click()
            time.sleep(1)
            print("Accepted cookies.")
        except NoSuchElementException:
            print("Cookie accept button not found.")

    def select_title(driver, title_text):
        title_dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '.bge-select__value-container'))
        )
        scroll_into_view(driver, title_dropdown)
        title_dropdown.click()
        script = f"""
        var options = document.querySelectorAll('.bge-select__option');
        for (var i = 0; i < options.length; i++) {{
            if (options[i].textContent.trim() === '{title_text}') {{
                options[i].click();
                return true;
            }}
        }}
        return false;
        """
        matched = driver.execute_script(script)
        if matched:
            print(f"Selected option: {title_text}")
        else:
            print(f"No match found for: {title_text}")

    def joint_select_title(driver, joint_title_text):
        try:
            joint_title_dropdown = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div[1]/div/section/div[2]/div[3]/div[1]/main/form/div[5]/div[1]/div/div'))
            )
            scroll_into_view(driver, joint_title_dropdown)
            joint_title_dropdown.click()
            script = f"""
            var options = document.querySelectorAll('.bge-select__option');
            for (var i = 0; i < options.length; i++) {{
                if (options[i].textContent.trim() === '{joint_title_text}') {{
                    options[i].click();
                    return true;
                }}
            }}
            return false;
            """
            matched = driver.execute_script(script)
            if matched:
                print(f"Selected joint title option: {joint_title_text}")
            else:
                print(f"No match found for joint title: {joint_title_text}")
        except Exception as e:
            print(f"Error selecting joint title: {e}")

    def nominated_select_title(driver, nominated_title_text):
        try:
            nominated_title_dropdown = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[@data-testid="nominated"]//div[contains(@class, "bge-select__value-container")]'))
            )
            scroll_into_view(driver, nominated_title_dropdown)
            nominated_title_dropdown.click()
            script = f"""
            var options = document.querySelectorAll('.bge-select__option');
            for (var i = 0; i < options.length; i++) {{
                if (options[i].textContent.trim() === '{nominated_title_text}') {{
                    options[i].click();
                    return true;
                }}
            }}
            return false;
            """
            matched = driver.execute_script(script)
            if matched:
                print(f"Selected nominated title option: {nominated_title_text}")
            else:
                print(f"No match found for nominated title: {nominated_title_text}")
        except Exception as e:
            print(f"Error selecting nominated title: {e}")

    def fill_eircode(driver, eircode):
        try:
            eircode_field = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div[1]/div/section/div[2]/div[3]/div[1]/main/form/div[7]/div/div'))
            )
            scroll_into_view(driver, eircode_field)
            driver.execute_script("arguments[0].value = arguments[1];", eircode_field, eircode)
            driver.execute_script("""
                var event = new Event('input', { bubbles: true });
                arguments[0].dispatchEvent(event);
                event = new Event('change', { bubbles: true });
                arguments[0].dispatchEvent(event);
            """, eircode_field)
            print(f"Successfully filled EIRCODE field with value: {eircode}")
        except Exception as e:
            print(f"Failed to fill EIRCODE field: {e}")
    

    def fill_text_field(driver, field_name, text):
        try:
            field = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f'//input[@name="{field_name}"]'))
            )
            scroll_into_view(driver, field)
            field.clear()
            field.send_keys(text)
        except Exception as e:
            print(f"Failed to fill field {field_name}: {e}")

    def select_phone_type(driver, phone_type_text):
        phone_dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div[1]/div/section/div[2]/div[3]/div[1]/main/form/div[1]/div[7]/div/div[1]/div/div/div'))
        )
        scroll_into_view(driver, phone_dropdown)
        phone_dropdown.click()
        driver.execute_script(f"""
        var options = document.querySelectorAll('.bge-select__option');
        for (var i = 0; i < options.length; i++) {{
            if (options[i].textContent.trim() === '{phone_type_text}') {{
                options[i].click();
                break;
            }}
        }}
        """)

    def joint_select_phone_type(driver, joint_phone_type_text):
        phone_dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div[1]/div/section/div[2]/div[3]/div[1]/main/form/div[5]/div[6]/div[1]/div/div/div'))
        )
        scroll_into_view(driver, phone_dropdown)
        phone_dropdown.click()
        driver.execute_script(f"""
        var options = document.querySelectorAll('.bge-select__option');
        for (var i = 0; i < options.length; i++) {{
            if (options[i].textContent.trim() === '{joint_phone_type_text}') {{
                options[i].click();
                break;
            }}
        }}
        """)

    def select_contact_type(driver, contact_type_text):
        phone_dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div[1]/div/section/div[2]/div[3]/div[1]/main/form/div[1]/div[7]/div/div[2]/div/div/div'))
        )
        scroll_into_view(driver, phone_dropdown)
        phone_dropdown.click()
        driver.execute_script(f"""
        var options = document.querySelectorAll('.bge-select__option');
        for (var i = 0; i < options.length; i++) {{
            if (options[i].textContent.trim() === '{contact_type_text}') {{
                options[i].click();
                break;
            }}
        }}
        """)

    def joint_select_contact_type(driver, joint_contact_type_text):
        phone_dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div[1]/div/section/div[2]/div[3]/div[1]/main/form/div[5]/div[6]/div[2]/div/div/div'))
        )
        scroll_into_view(driver, phone_dropdown)
        phone_dropdown.click()
        driver.execute_script(f"""
        var options = document.querySelectorAll('.bge-select__option');
        for (var i = 0; i < options.length; i++) {{
            if (options[i].textContent.trim() === '{joint_contact_type_text}') {{
                options[i].click();
                break;
            }}
        }}
        """)

    def manipulate_and_click_button(driver, xpath):
        try:
            next_button = WebDriverWait(driver, 20).until(
                EC.visibility_of_element_located((By.XPATH, xpath))
            )
            scroll_into_view(driver, next_button)
            next_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            print("Next button is clickable.")
            next_button.click()
            print("The Next button was clicked.")
        except Exception as e:
            print(f"Error clicking the next button: {e}")
            driver.save_screenshot('error_screenshot.png')

    def click_element_with_js(driver, xpath):
        try:
            element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            scroll_into_view(driver, element)
            driver.execute_script("arguments[0].click();", element)
            print(f"Clicked element with XPath: {xpath}")
        except Exception as e:
            print(f"Error clicking element with XPath: {xpath}: {e}")

    def datetopay_select_title(driver, date_text):
        try:
            datetopay_title_dropdown = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div[1]/div/section/div[2]/div[3]/div[1]/main/form/div[3]/div[2]/div/div/div/div[1]'))
            )
            scroll_into_view(driver, datetopay_title_dropdown)
            datetopay_title_dropdown.click()
            script = f"""
            var options = document.querySelectorAll('.bge-select__option');
            for (var i = 0; i < options.length; i++) {{
                if (options[i].textContent.trim() === '{date_text}') {{
                    options[i].click();
                    return true;
                }}
            }}
            return false;
            """
            matched = driver.execute_script(script)
            if matched:
                print(f"Selected Datetopay title option: {date_text}")
            else:
                print(f"No match found for Datetopay title: {date_text}")
        except Exception as e:
            print(f"Error selecting Datetopay title: {e}")

    for index, row in df.iterrows():
        try:
            driver = initialize_driver()
            accept_cookies(driver)

            # Wait for the initial button to be clickable and click
            initialclick = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[1]/div/section/div[2]/div/div[3]/div[2]/footer/a"))
            )
            scroll_into_view(driver, initialclick)
            initialclick.click()

            if row['Customer Type'] == 'Yes':
                yesnec = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[1]/div/section/div/div/div[2]/div/div[1]/form/div/div[1]/div/div/div[2]/div[1]/label"))
                )
                scroll_into_view(driver, yesnec)
                yesnec.click()
            else:
                nonec = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[1]/div/section/div/div/div[2]/div/div[1]/form/div/div[1]/div/div/div[2]/div[2]/label"))
                )
                scroll_into_view(driver, nonec)
                nonec.click()

            if row['Category'] == 'Dual Fuel':
                dultg = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[1]/div/section/div/div/div[2]/div/div[1]/form/div/div[2]/div/div/div[2]/div[1]/label"))
                )
                scroll_into_view(driver, dultg)
                dultg.click()
            elif row['Category'] == 'Electricity':
                eleltg = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[1]/div/section/div/div/div[2]/div/div[1]/form/div/div[2]/div/div/div[2]/div[2]/label"))
                )
                scroll_into_view(driver, eleltg)
                eleltg.click()
            else:
                gas = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[1]/div/section/div/div/div[2]/div/div[1]/form/div/div[2]/div/div/div[2]/div[3]/label"))
                )
                scroll_into_view(driver, gas)
                gas.click()

            if not (row['Customer Type'] == 'Yes' and row['Category'] == 'Gas'):
                if row['Meter_type'] == 'Standard':
                    staemt = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[1]/div/section/div/div/div[2]/div/div[1]/form/div/div[3]/div/div/div/div[2]/div[1]/label"))
                    )
                    scroll_into_view(driver, staemt)
                    staemt.click()
                elif row['Meter_type'] == 'Smart':
                    smaemt = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[1]/div/section/div/div/div[2]/div/div[1]/form/div/div[3]/div/div/div/div[2]/div[2]/label"))
                    )
                    scroll_into_view(driver, smaemt)
                    smaemt.click()

            plan_name = row['Tariff Title']
            plan_found = False
            while not plan_found:
                try:
                    plan_xpath = plan_xpath_mapping[plan_name]
                    view_plan_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, plan_xpath))
                    )
                    scroll_into_view(driver, view_plan_button)
                    view_plan_button.click()
                    plan_found = True
                except Exception as e:
                    try:
                        show_more_button = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div[1]/div/section/div/div/div[2]/div/div[2]/div[2]/button'))
                        )
                        scroll_into_view(driver, show_more_button)
                        show_more_button.click()
                        time.sleep(2)
                    except:
                        print(f"Plan '{plan_name}' not found.")
                        break

            if plan_found:
                try:
                    switch_now_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, '//button[contains(@class, "button primary-sub arrow") and text()="Switch now"]'))
                    )
                    scroll_into_view(driver, switch_now_button)
                    switch_now_button.click()
                    try:
                        continue_button = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div[2]/div/div/div/section/div/div/div[2]/button[2]'))
                        )
                        scroll_into_view(driver, continue_button)
                        continue_button.click()
                    except:
                        pass

                    title_text = row['Salutation']
                    select_title(driver, title_text)

                    fill_text_field(driver, 'mainDetails.firstName', row['First Name'])
                    fill_text_field(driver, 'mainDetails.lastName', row['Surname'])
                    fill_text_field(driver, 'mainDetails.email', row['Email Address'])
                    fill_text_field(driver, 'mainDetails.emailConfirmation', row['Confirm Email Address'])

                    select_phone_type(driver, row['Phone Type'])
                    select_contact_type(driver, row['Contact type'])
                    fill_text_field(driver, 'mainDetails.phoneNumber', row['Mobile Number'])

                    if row['Occupancy Status'] == 'Homeowner':
                        homeowner_label = driver.find_element(By.CSS_SELECTOR, "label[for='Homeowner']")
                        scroll_into_view(driver, homeowner_label)
                        homeowner_label.click()
                    elif row['Occupancy Status'] == 'Landlord':
                        landlord_label = driver.find_element(By.CSS_SELECTOR, "label[for='Landlord']")
                        scroll_into_view(driver, landlord_label)
                        landlord_label.click()
                    elif row['Occupancy Status'] == 'Tenant':
                        tenant_label = driver.find_element(By.CSS_SELECTOR, "label[for='Tenant']")
                        scroll_into_view(driver, tenant_label)
                        tenant_label.click()

                    if row['Third Party'] == 'Yes':
                        third_party = driver.find_element(By.CSS_SELECTOR, "label[for='thirdPartyOrder']")
                        scroll_into_view(driver, third_party)
                        third_party.click()

                    if row['Joint Account'] == 'Yes':
                        try:
                            joint_account = driver.find_element(By.CSS_SELECTOR, "label[for='jointAccountInfo.jointAccount']")
                            scroll_into_view(driver, joint_account)
                            joint_account.click()
                            joint_title_text = row['joint_account_holder_title']
                            joint_select_title(driver, joint_title_text)
                            fill_text_field(driver, 'jointAccountInfo.firstName', row['joint_account_holder_first_name'])
                            fill_text_field(driver, 'jointAccountInfo.lastName', row['joint_account_holder_last_name'])
                            fill_text_field(driver, 'jointAccountInfo.email', row['Joint Account Email'])
                            fill_text_field(driver, 'jointAccountInfo.emailConfirmation', row['Joint Account Confirm Email'])
                            joint_select_phone_type(driver, row['Joint Phone Type'])
                            joint_select_contact_type(driver, row['Joint Contact Type'])
                            contact_number = row['joint_account_holder_ Mobile Number']
                            fill_text_field(driver, 'jointAccountInfo.phoneNumber', contact_number)
                        except Exception as e:
                            print(f"Error handling joint account: {e}")

                    if row['Is Home Address same as Billing Address?'] == 'Yes':
                        try:
                            different_address = driver.find_element(By.CSS_SELECTOR, "label[for='billingAddress.billAddress']")
                            scroll_into_view(driver, different_address)
                            different_address.click()
                            fill_eircode(driver, row['Different Address EIRCODE'])
                        except Exception as e:
                            print(f"Error handling Different Address: {e}")

                    if row['Nominated User'] == 'Yes':
                        try:
                            nominated_user = driver.find_element(By.CSS_SELECTOR, "label[for='nominated']")
                            scroll_into_view(driver, nominated_user)
                            nominated_user.click()
                            time.sleep(1)
                            nominated_title_text = row['Nominated Title']
                            nominated_select_title(driver, nominated_title_text)
                            fill_text_field(driver, 'nominatedPerson.firstName', row['Nominated Fname'])
                            fill_text_field(driver, 'nominatedPerson.lastName', row['Nominated LName'])
                            driver.execute_script("window.scrollBy(0, 150);")
                            print("Scrolled a little below after filling nominated user details.")
                            time.sleep(1)
                            next_button = WebDriverWait(driver, 20).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.button.primary.arrow.inline'))
                            )
                            scroll_into_view(driver, next_button)
                            print("Next button is clickable.")
                            next_button.click()
                            try:
                                popup_button = WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[2]/div/div/div/footer/div/button[2]"))
                                )
                                scroll_into_view(driver, popup_button)
                                popup_button.click()
                                print("Clicked 'Yes, it's right' button.")
                            except:
                                pass
                        except:
                            print("No nominations")

                    if row['Category'] == 'Gas':
                        fill_text_field(driver, 'gasDetails.gprn', row['GPRN'])
                        next_button_xpath = '/html/body/div[1]/div/div[1]/div/section/div[2]/div[3]/div[1]/main/form/main/div/div/div[2]/div/button'
                        manipulate_and_click_button(driver, next_button_xpath)
                        next_button_xpath = '/html/body/div[1]/div/div[2]/div/div/div/footer/div/button[2]'
                        manipulate_and_click_button(driver, next_button_xpath)
                        fill_text_field(driver, 'gasDetails.gasMeterReading', row['Gas Meter Reading'])

                    elif row['Category'] == 'Electricity' and row['Meter_type'] == 'Smart':
                        fill_text_field(driver, 'electricityDetails.mprn', row['MPRN'])
                        next_button_xpath = '/html/body/div[1]/div/div[2]/div/div/div/footer/div/button[2]'
                        manipulate_and_click_button(driver, next_button_xpath)
                    elif row['Category'] == 'Electricity' and row['Meter_type'] == 'Standard':
                        fill_text_field(driver, 'electricityDetails.mprn', row['MPRN'])
                        next_button_xpath = '/html/body/div[1]/div/div[2]/div/div/div/footer/div/button[2]'
                        manipulate_and_click_button(driver, next_button_xpath)

                    elif row['Category'] == 'Dual Fuel' and row['Meter_type'] == 'Smart':
                        fill_text_field(driver, 'gasDetails.gprn', row['GPRN'])
                        next_button_xpath = '/html/body/div[1]/div/div[1]/div/section/div[2]/div[3]/div[1]/main/form/main/div/div/div[2]/div/button'
                        manipulate_and_click_button(driver, next_button_xpath)
                        next_button_xpath = '/html/body/div[1]/div/div[2]/div/div/div/footer/div/button[2]'
                        manipulate_and_click_button(driver, next_button_xpath)
                        fill_text_field(driver, 'gasDetails.gasMeterReading', row['Gas Meter Reading'])
                        fill_text_field(driver, 'electricityDetails.mprn', row['MPRN'])
                        next_button_xpath = '/html/body/div[1]/div/div[2]/div/div/div/footer/div/button[2]'
                        manipulate_and_click_button(driver, next_button_xpath)
                    elif row['Category'] == 'Dual Fuel' and row['Meter_type'] == 'Standard':
                        fill_text_field(driver, 'gasDetails.gprn', row['GPRN'])
                        next_button_xpath = '/html/body/div[1]/div/div[1]/div/section/div[2]/div[3]/div[1]/main/form/main/div/div/div[2]/div/button'
                        manipulate_and_click_button(driver, next_button_xpath)
                        next_button_xpath = '/html/body/div[1]/div/div[2]/div/div/div/footer/div/button[2]'
                        manipulate_and_click_button(driver, next_button_xpath)
                        fill_text_field(driver, 'gasDetails.gasMeterReading', row['Gas Meter Reading'])
                        fill_text_field(driver, 'electricityDetails.mprn', row['MPRN'])
                        next_button_xpath = '/html/body/div[1]/div/div[2]/div/div/div/footer/div/button[2]'
                        manipulate_and_click_button(driver, next_button_xpath)
                    else:
                        print("No provider Found")

                    confirmation_checkbox_xpath = "//input[@id='supplyAddress.openAccountAuthorization']"
                    if row['Confirm Authorisation'] == 'Yes':
                        click_element_with_js(driver, confirmation_checkbox_xpath)
                    else:
                        pass

                    if row['Movein'] == 'Yes':
                        move_in_yes_radio_xpath = "//input[@id='true']"
                        click_element_with_js(driver, move_in_yes_radio_xpath)
                    else:
                        move_in_no_radio_xpath = "//input[@id='false']"
                        click_element_with_js(driver, move_in_no_radio_xpath)

                    vulnerablecustomers_checkbox_xpath = "//input[@id='specialServicesIndicator']"
                    if row['vulnerable_customers'] == 'Yes':
                        click_element_with_js(driver, vulnerablecustomers_checkbox_xpath)
                    else:
                        pass

                    Support_Type_checkbox_xpath = "//input[@id='priorityCustomerIndicator']"
                    if row['Is On Life Support?'] == 'Yes':
                        click_element_with_js(driver, Support_Type_checkbox_xpath)
                    else:
                        pass

                    next_button_xpath1 = '/html/body/div/div/div[1]/div/section/div[2]/div[3]/div[1]/main/form/div[5]/button[2]'
                    manipulate_and_click_button(driver, next_button_xpath1)

                    fill_text_field(driver, 'paymentDetails.directDebitAccountName', row['Bank account holder name'])
                    fill_text_field(driver, 'paymentDetails.iban', row['IBAN'])

                    standard_smart_plans = [
                        "Green Standard Smart Dual Fuel and €100 Welcome Bonus",
                        "Green Standard Smart Electricity and €50 Welcome Bonus",
                        "Green Standard Smart Electricity Discount + €50 Welcome Bonus",
                        "Standard Smart Dual Fuel and €100 Welcome Bonus",
                        "Standard Smart Dual Fuel and Free Hive",
                        "Standard Smart Dual Fuel Free Hive+Amazon Echo Dot",
                        "Standard Smart Electricity Discount",
                        "Standard Smart Electricity Discount + €50 Welcome Bonus",
                        "Standard Smart Electricity Free Hive + Free Amazon Echo Dot",
                        "Green Standard Smart Dual Fuel Discount",
                        "Green Standard Smart Electricity Discount",
                        "Standard Smart Duel Fuel Discount",
                        "Standard Smart Dual Fuel Discount",
                        "Standard Smart Electricity with Free Hive + Free Amazon Echo Dot"
                    ]
                    if row['Tariff Title'] in standard_smart_plans:
                        pass
                    else:
                        try:
                            if row['Billing Type'] == 'Monthly':
                                Billed_Monthly_xpath = "//input[@id='Monthly']"
                                click_element_with_js(driver, Billed_Monthly_xpath)
                            else:
                                Billed_Bimonthly_xpath = "//input[@id='Bimonthly']"
                                click_element_with_js(driver, Billed_Bimonthly_xpath)
                        except:
                            pass
                        time.sleep(1)
                        try:
                            date_text = row['Preferred Day of Billing']
                            datetopay_select_title(driver, date_text)
                        except:
                            pass

                    try:
                        acceptdebits_checkbox_xpath = "//input[@id='paymentDetails.directDebitAuthorization']"
                        if row['Accept Debits'] == 'Yes':
                            click_element_with_js(driver, acceptdebits_checkbox_xpath)
                    except:
                        pass

                    try:
                        termsandconditions_checkbox_xpath = "//input[@id='paymentDetails.termsAndConditionsAccepted']"
                        if row['Terms Conditions'] == 'Yes':
                            click_element_with_js(driver, termsandconditions_checkbox_xpath)
                    except:
                        pass

                    try:
                        next_button_xpath2 = "//button[@type='submit' and contains(@class, 'button primary arrow inline') and text()='Next']"
                        manipulate_and_click_button(driver, next_button_xpath2)
                    except Exception as e:
                        print(f"Next button not found: {e}")

                    try:
                        RighttoCancel_checkbox_xpath = "//input[@id='rightToCancel']"
                        if row['Right to Cancel'] == 'Yes':
                            click_element_with_js(driver, RighttoCancel_checkbox_xpath)
                    except Exception as e:
                        print(f"Right to Cancel field not found: {e}")

                    try:
                        configurepermenetly_checkbox_xpath = "//input[@id='smartMeterConfigure']"
                        if row['Configure Smart Meter Permanently'] == 'Yes':
                            click_element_with_js(driver, configurepermenetly_checkbox_xpath)
                    except Exception as e:
                        print(f"Configure smart meter permanently field not found: {e}")

                    try:
                        Tandcsupply_checkbox_xpath = "//input[@id='termsAndConditions']"
                        if row['Terms and conditions supply of electricity'] == 'Yes':
                            click_element_with_js(driver, Tandcsupply_checkbox_xpath)
                    except Exception as e:
                        print(f"Terms and conditions supply of electricity field not found: {e}")
                    
                    # try:
                    #     next_button_xpath3 = "/html/body/div/div/div[1]/div/section/div[2]/div[3]/div[1]/main/form/div[2]/button[2]"
                    #     manipulate_and_click_button(driver, next_button_xpath3)
                    # except Exception as e:
                    #     print(f"Next button not found: {e}")
                        
                        
                    

                    print("Successfully Completed the process")
                    success_data.append({
                        'MPRN': row['MPRN'],
                        'First Name': row['First Name'],
                        'Status': 'Success'
                    })
                    print(f"Success: MPRN {row['MPRN']}, Name: {row['First Name']}")
                except:
                    print("Plan not found")
        except Exception as e:
            print(f"Error found in row {index + 1}: {e}")
            failed_data.append({
                'MPRN': row['MPRN'],
                'First Name': row['First Name'],
                'Status': 'Failed',
                'Error': str(e)
            })
            print(f"Failed: MPRN {row['MPRN']}, Name: {row['First Name']} - Error: {str(e)}")
        finally:
            time.sleep(2)
            driver.quit()

    # output_file = "output.xlsx"
    # output_path = os.path.join(settings.MEDIA_ROOT, output_file)
    # if os.path.exists(output_path):
    #     os.remove(output_path)
    #     df_success = pd.DataFrame(success_data)
    #     df_failed = pd.DataFrame(failed_rows)
    # with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    #     df_success.to_excel(writer, sheet_name='Passed', index=False)
    #     df_failed.to_excel(writer, sheet_name='Failed', index=False)
        
    print("\nSummary:")
    print("Successful Rows:")
    for item in success_data:
        print(f"MPRN: {item['MPRN']}, Name: {item['First Name']}")

    print("\nFailed Rows:")
    for item in failed_data:
        print(f"MPRN: {item['MPRN']}, Name: {item['First Name']}, Error: {item['Error']}")
