import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter as tk
import tkinter.font as tkFont
import threading
from selenium.webdriver.chrome.options import Options

def fetch_pdp_Makro(driver, url,check_once):
            model_id = ""

            
            try:
                if check_once == 0:
                    driver.execute_script("window.open('');")
                    check_once+=1
                driver.switch_to.window(driver.window_handles[1])
                driver.get(url)
                time.sleep(10)
                # driver.set_page_load_timeout(30)
                
                try:
                    ids = driver.find_element(By.XPATH,"//*[contains(text(), 'Model')]")
                    parent = ids.find_element(By.XPATH, '..')
                    children = parent.find_elements(By.XPATH, '*')
                    skip_first = False
                    for child in children:
                        if skip_first == False:
                             skip_first = True
                        else:
                            model_id=child.text
                            break
                    skip_first = False
                    # print(model_id)
                except:
                    # driver.close()
                    # driver.switch_to.window(driver.window_handles[0])
                    model_id = ""
                

        


                
                

                # driver.close()
                driver.switch_to.window(driver.window_handles[0])
                return model_id,check_once
            except:
                return model_id,check_once
    


def InfiniteScrolling(driver):
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            # Scroll down to bottom
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

            # Wait to load page
            time.sleep(4)

            # Calculate new scroll height and compare with last scroll height
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height



def Hirsch_Web(driver,list_of_categories,data,Sharaf_DG):
        output_df = pd.DataFrame(columns=['Model','Hirsch'])
        for cate in list_of_categories:
            df = data[data['Category'] == cate]
            # print(df)
            list_of_models = df["Models"]
            # We got the models of one category
            # check_once = 0
            for models in list_of_models:
                driver.get("https://www.hirschs.co.za/search/"+models)
                # time.sleep(5)
                try:
                    ids = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".toolbar-number"))).text
                    if (int(ids)>0):

                        product_link = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".item.product.product-item")))
                        product_name = product_link.find_element(By.CSS_SELECTOR, ".product-item-link").text
                        link = product_link.find_element(By.CSS_SELECTOR, ".product-item-link").get_attribute("href")
                        
                        if product_name.find(models) != -1:
                            output_df = output_df.append({
                                    "Model":models,
                                    "Hirsch": "o",
                                    "Product Link": link
                            },ignore_index=True)
                            print(models,"Found")
                        else:
                            output_df = output_df.append({
                                "Model":models,
                                "Hirsch": "x",
                                "Product Link": ""
                            },ignore_index=True)
                            print(models,"Not Found")
                    else:
                        output_df = output_df.append({
                            "Model":models,
                            "Hirsch": "x",
                            "Product Link": ""
        
                        },ignore_index=True)
                        print(models,"Not Found")
                except:
                    output_df = output_df.append({
                            "Model":models,
                            "Hirsch": "x",
                            "Product Link": ""
      
                    },ignore_index=True)
                    print(models,"Not Found")

                
                # if check_once == 0:
                #     print("Model: ",models)
                #     df_link = Sharaf_DG[Sharaf_DG['Category'] == cate]
                #     keyword = models
                #     # print(df_link["Links"])
                #     dyno_link = df_link["Links"].iloc[0]
                #     # print(dyno_link)
           
                #     model_ids :list = []
                #     driver.get(dyno_link)
                #     # # Get scroll height
                #     # InfiniteScrolling(driver)
                    
                #     # driver.get(dyno_link)
                #     time.sleep(5)
                #     ids = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".products.list.items.product-items")))
                #     all_divs  = ids.find_elements(By.CSS_SELECTOR, ".item.product.product-item")
                #     # print(len(all_divs))
                #     # Compare product name with model name 
                #     for div in all_divs:
                #         title = div.find_element(By.CSS_SELECTOR,".product-item-link")
                #         title_value = title.text
                #         model_id = title_value
                #         print(model_id)
                #         check_once = 1
                #         # Save this model id in the list and use it later 
                #         # 
                #         model_ids.append(model_id)



                # total_models = len(model_ids)
                # counter = 0
                # for each_model in model_ids: 
                #     if each_model.find(models) != -1:
                #         output_df = output_df.append({
                #                 "Model":models,
                #                 "Hirsch": "o",
                #         },ignore_index=True)
                #         print(models,"Found")
                #         break
                #     counter+=1

                # if counter == total_models:
                #     output_df = output_df.append({
                #             "Model":models,
                #             "Hirsch": "x",
      
                #     },ignore_index=True)
                #     print(models,"Not Found")
                
                # Compare here now

        with pd.ExcelWriter("output.xlsx",mode="a",if_sheet_exists='replace') as writer:
            output_df.to_excel(writer,sheet_name="Hirsch")
        
def Makro_Web(driver,list_of_categories,data,Sharaf_DG):
        output_df = pd.DataFrame(columns=['Model','Makro'])
        for cate in list_of_categories:
            df = data[data['Category'] == cate]
            # print(df)
            list_of_models = df["Models"]
            # We got the models of one category
            check_once = 0
            for models in list_of_models:
                if check_once == 0:
                    print("Model: ",models)
                    df_link = Sharaf_DG[Sharaf_DG['Category'] == cate]
                    keyword = models
                    # print(df_link["Links"])
                    dyno_link = df_link["Links"].iloc[0]
                    # print(dyno_link)
           
                    model_ids :list = []
                    driver.get(dyno_link)
                    # # Get scroll height
                    # InfiniteScrolling(driver)
                    
                    # driver.get(dyno_link)
                    time.sleep(10)
                    try:
                        ids = driver.find_elements(By.CSS_SELECTOR,".css-1dbjc4n.r-1m04atk.r-wk8lta")

                        count =0 
                        for eid in ids:
                            if count ==1:
                                all_divs  = eid.find_elements(By.CSS_SELECTOR, ".css-901oao.css-cens5h.r-1qimiim.r-ukjfwf.r-ubezar.r-135wba7.r-c8eef1")
                            count +=1
                        
                        if len(ids) <=1:
                            ids = driver.find_element(By.CSS_SELECTOR,".css-1dbjc4n.r-1m04atk.r-wk8lta")
                            all_divs  = ids.find_elements(By.CSS_SELECTOR, ".css-901oao.css-cens5h.r-1qimiim.r-ukjfwf.r-ubezar.r-135wba7.r-c8eef1")
                    except:
                        all_divs = driver.find_elements(By.CSS_SELECTOR,".product-tile-inner__productTitle.js-gtmProductLinkClickEvent.text-overflow-ellipsis.line-clamp-2")
                    counter = 0
                    print(keyword)
                    for div in all_divs:
                        link = div.find_element(By.TAG_NAME,"a").get_attribute("href")
                        print("Link:",link)
                        model_id,check_once = fetch_pdp_Makro(driver,link,check_once)
                        print(model_id)
                        check_once = 1
                        # Save this model id in the list and use it later 
                        # 
                        model_ids.append(model_id)



                total_models = len(model_ids)
                counter = 0
                for each_model in model_ids: 
                    if each_model.find(models) != -1:
                        output_df = output_df.append({
                                "Model":models,
                                "Makro": "o",
                        },ignore_index=True)
                        print(models,"Found")
                        break
                    counter+=1

                if counter == total_models:
                    output_df = output_df.append({
                            "Model":models,
                            "Makro": "x",
      
                    },ignore_index=True)
                    print(models,"Not Found")
                
                # Compare here now

        with pd.ExcelWriter("output.xlsx",mode="a",if_sheet_exists='replace') as writer:
            output_df.to_excel(writer,sheet_name="Makro")
 

def Run_Hirsch():




    # 0------------------------------------------------------------------------------
    # df_sharaf_dg_categories_keywords = pd.read_excel("search_keywords.xlsx")
    # df_sharaf_dg_brands = pd.read_excel("input.xlsx")
    data = pd.read_excel("models.xlsx",sheet_name="Models")
    Sharaf_DG = pd.read_excel("models.xlsx",sheet_name="Hirsch")
    # LULU = pd.read_excel("models.xlsx",sheet_name="LULU")
    # Jumbo = pd.read_excel("models.xlsx",sheet_name="Jumbo")
    # output_df = pd.DataFrame(columns=['Model','Sharaf_DG'])
    # print(data["Model"])
    
    # print(data["Category"].unique())
    # driver = webdriver.Chrome(options=chrome_options)
    driver = webdriver.Chrome("C:\Program Files\chromedriver.exe")
    list_of_categories = data["Category"].unique()

    Hirsch_Web(driver,list_of_categories,data,Sharaf_DG)
    
  

def Run_Makro():




    # 0------------------------------------------------------------------------------
    # df_sharaf_dg_categories_keywords = pd.read_excel("search_keywords.xlsx")
    # df_sharaf_dg_brands = pd.read_excel("input.xlsx")
    data = pd.read_excel("models.xlsx",sheet_name="Models")
    Sharaf_DG = pd.read_excel("models.xlsx",sheet_name="Makro")
    # LULU = pd.read_excel("models.xlsx",sheet_name="LULU")
    # Jumbo = pd.read_excel("models.xlsx",sheet_name="Jumbo")
    # output_df = pd.DataFrame(columns=['Model','Sharaf_DG'])
    # print(data["Model"])
    
    # print(data["Category"].unique())
    # driver = webdriver.Chrome(options=chrome_options)
    driver = webdriver.Chrome("C:\Program Files\chromedriver.exe")
    list_of_categories = data["Category"].unique()

    Makro_Web(driver,list_of_categories,data,Sharaf_DG)
 


# Main App 
class App:

    def __init__(self, root):
        #setting title
        root.title("South Africa Model Check")
        ft = tkFont.Font(family='Arial Narrow',size=13)
        #setting window size
        width=640
        height=480
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)
        root.configure(bg='black')

        ClickBtnLabel=tk.Label(root)
       
      
        
        ClickBtnLabel["font"] = ft
        
        ClickBtnLabel["justify"] = "center"
        ClickBtnLabel["text"] = "South Africa Model Check"
        ClickBtnLabel["bg"] = "black"
        ClickBtnLabel["fg"] = "white"
        ClickBtnLabel.place(x=120,y=190,width=150,height=70)
    

        
        Lulu=tk.Button(root)
        Lulu["anchor"] = "center"
        Lulu["bg"] = "#009841"
        Lulu["borderwidth"] = "0px"
        
        Lulu["font"] = ft
        Lulu["fg"] = "#ffffff"
        Lulu["justify"] = "center"
        Lulu["text"] = "START"
        Lulu["relief"] = "raised"
        Lulu.place(x=375,y=190,width=150,height=70)
        Lulu["command"] = self.start_func




  

    def ClickRun(self):

        running_actions = [
            Run_Hirsch,          
            # Run_Makro,
            # Run_Jumbo
        ]

        thread_list = [threading.Thread(target=func) for func in running_actions]

        # start all the threads
        for thread in thread_list:
            thread.start()

        # wait for all the threads to complete
        for thread in thread_list:
            thread.join()
    
    def start_func(self):
        thread = threading.Thread(target=self.ClickRun)
        thread.start()

    
        

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()


# Run()
