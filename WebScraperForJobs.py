import asyncio
from pyppeteer import launch
import time
import smtplib
from email.message import EmailMessage
import win32com.client as win32


class WebScraper:

    async def main():
        await WebScraper.skim()

    @staticmethod
    async def skim():
        browser = await launch(executablePath=r'C:\Program Files\Google\Chrome\Application\chrome.exe', headless=False)

        page = await browser.newPage()
        print("Navigating page")
        
        await page.goto('https://www.indeed.com')
        await page.waitForSelector('#text-input-what')
        await page.waitForSelector('#text-input-where')
        await page.type('#text-input-what', 'Internship')
        await page.type('#text-input-where', '')
        await page.click('button[type="submit"]')
        await page.waitForNavigation()

        job_listings = await page.querySelectorAll('.resultContent')
        jobs = []
        for job in job_listings:
            # Extract the job title
            title_element = await job.querySelector('h2.jobTitle span[title]')
            title = await page.evaluate('(element) => element.textContent', title_element)
            hyperlink_element = await job.querySelector('a')
            hyperlink = await page.evaluate('(element) => element.href', hyperlink_element)

            # Extract the company name
            company_element = await job.querySelector('div.company_location [data-testid="company-name"]')
            company = await page.evaluate('(element) => element.textContent', company_element)


            # Extract the location
            location_element = await job.querySelector('div.company_location [data-testid="text-location"]')
            location = await page.evaluate('(element) => element.textContent', location_element)
            if "intern" in title.lower() and ("CT" in location or "Connecticut" in location) and "2025" in title.lower():
                jobs.append([title, company, location, hyperlink])
                #print({'title': title, 'company': company, 'location': location, 'hyperlink': hyperlink})
        

        print("\nFormatted job details:")
        for job in jobs:
            print(f"Title: {job[0]}")
            print(f"Company: {job[1]}")
            print(f"Location: {job[2]}")
            print(f"Hyperlink: {job[3]}")
            print("-" * 40)  # Separator between jobs
        
        # Print each job's details in a more readable format
        bodyOfEmail = ""
        for job in jobs:
            bodyOfEmail += f"<p><strong>I'm working on a web skimmer that skims for jobs and emails them to me! Check it out.</strong><br>"
            bodyOfEmail += f"<p><strong>Title:</strong> {job[0]}<br>"
            bodyOfEmail += f"<strong>Company:</strong> {job[1]}<br>"
            bodyOfEmail += f"<strong>Location:</strong> {job[2]}<br>"
            bodyOfEmail += f"<strong>Hyperlink:</strong> <a href='{job[3]}'>{job[3]}</a></p>"
            bodyOfEmail += "<hr>"  # Adds a horizontal line to separate jobs
    
        print(bodyOfEmail)
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        #mail.To = 'rex.dorchester@proton.me; emilia.j913@gmail.com'
        mail.To = 'rex.dorchester@proton.me; arena.alex89@gmail.com'
        
        mail.Subject = 'Custom Job Query'
        mail.Body = bodyOfEmail
        mail.HTMLBody = bodyOfEmail #this field is optional

        mail.Send()


if __name__ == "__main__":
    asyncio.run(WebScraper.main())