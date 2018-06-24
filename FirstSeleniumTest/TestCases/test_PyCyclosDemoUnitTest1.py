import time
import unittest


from selenium import webdriver


class CyclosDemoUnitTest1(unittest.TestCase):

    def setUp(self):
        self.driver = webdriver.Remote(
            command_executor="http://localhost:4444/wd/hub",
            desired_capabilities={
                "browserName": "chrome",
            })

        self.driver.implicitly_wait(30)
        self.driver.maximize_window()

    def test_PyCyclosDemoUnitTest1(self):
        driver = self.driver
        driver.get("https://demo.cyclos.org/#home")
        driver.set_page_load_timeout(10)
        driver.maximize_window()
        self.assertIn("Cyclos",driver.title)
        time.sleep(5)
        elem = driver.find_element_by_xpath("/html/body/div[3]/div[1]/div/ul/li[2]/a/span")
        elem.click()
        time.sleep(4)
        elem = driver.find_element_by_xpath(
            "/html/body/div[3]/div[3]/div/div/div[2]/div[2]/div[2]/div/div/div/div[2]/form/div/input")
        elem.send_keys("demo")

        elem = driver.find_element_by_xpath("//input[@type='password']")
        elem.send_keys("1234")
        time.sleep(2)
        elem = driver.find_element_by_xpath("//div[@class='actionButtonText']")
        elem.click()
        time.sleep(4)
        assert "Fail to login" not in driver.page_source
        time.sleep(4)
        elem = driver.find_element_by_xpath("//span[@class='gwt-InlineLabel'][contains(text(),'Users')]")
        elem.click()
        time.sleep(4)
        elem = driver.find_element_by_xpath("//html//td[@class='formField firstFormField']/input[1]")
        elem.send_keys("nilesh")
        time.sleep(2)
        elem = driver.find_element_by_xpath("//div[@class='actionButtonText'][contains(text(),'Search')]")
        elem.click()
        time.sleep(4)
        elem = driver.find_element_by_xpath("//div[@class='tiledResultLabel']")
        elem.click()
        time.sleep(5)

        # User Page
        elem = driver.find_element_by_xpath("//span[@class='gwt-InlineLabel'][contains(text(),'Make payment')]")
        elem.click()
        time.sleep(4)
        elem = driver.find_element_by_xpath("//html//td[1]/input[1]")
        elem.send_keys("100")
        elem = driver.find_element_by_xpath("//textarea[@class='inputField full']")
        elem.send_keys("Money added to nilesh account using Cyclos App!!![Test Case 0]")
        elem = driver.find_element_by_xpath("//div[@class='actionButtonText']")
        elem.click()
        time.sleep(3)
        assert "Money added to nilesh account using Cyclos App!!![Test Case 0]" in driver.page_source
        elem = driver.find_element_by_xpath("//div[@class='actionButtonText']")
        elem.click()
        time.sleep(4)
        assert "The payment was successful" in driver.page_source
        print("Payment done !!!")
        elem = driver.find_element_by_xpath("//div[@class='actionButtonText'][contains(text(),'Close')]")
        elem.click()
        time.sleep(3)
        elem = driver.find_element_by_xpath("//span[@class='gwt-InlineLabel'][contains(text(),'Home')]")
        elem.click()
        time.sleep(2)
        elem = driver.find_element_by_xpath("//span[@class='statusMenuText'][contains(text(),'Logout')]")
        elem.click()
        print("Cyclos App logged out successfully !!!")

    def tearDown(self):
        self.driver.close()
        self.driver.quit()


if __name__ == "__main__":
    unittest.main()


