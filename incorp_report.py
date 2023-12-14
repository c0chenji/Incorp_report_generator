import pandas as pd

class ReportProcessor:
  def __init__(self,path,start_date, end_date):
    self.path = path
    self.start_date = start_date
    self.end_date = end_date
    self.main_df = None
    self.grand_total_df = None
    
  def read_and_clean_data(self):
    initial_df = pd.read_excel(self.path)
    
    initial_df["日期"] = pd.to_datetime(initial_df["日期"],format='ISO8601')

    sorted_df = initial_df[(initial_df['日期'] >= self.start_date) & (initial_df['日期'] <= self.end_date)]
    
    cleaned_df = sorted_df.sort_values(by='日期').drop(axis='columns',columns=["CAD","账户","货币", "账户.1","二级类别"])
    
    excluded_catogories_lst = ["超市/便利店"]
    
    cleaned_df = cleaned_df[~cleaned_df['分类'].isin(excluded_catogories_lst)]
    
    self.main_df = cleaned_df
    
  def get_desired_df(self):

    bank_account_fee_df = self.main_df[self.main_df['备注']=="Monthly account fee"]
    bank_account_fee_df = bank_account_fee_df.replace("其他","Bank Fee")

    travel_df = self.main_df[self.main_df['备注']=="CampAdelaide"]
    travel_df = travel_df.replace("其他","traveling")
    
    water_cost_df = self.main_df[self.main_df['备注']=="水"]
    
    hydro_cost_df = self.main_df[self.main_df['备注']=="电"]
    
    gas_cost_df = self.main_df[self.main_df['备注']=="气"]
    
    telephone_line_cost_df = self.main_df[self.main_df['备注']=="Lucky"]
  
    internet_df = self.main_df[self.main_df['备注']=="Bell Internet"]
    internet_df = internet_df.replace("辦公用品","Internet")
    
    parking_df = self.main_df[self.main_df['备注'].isin(["Honk Parking","Parking"])]
    
    fuel_df = self.main_df[self.main_df['备注'].isin(["Esso"])]
    
    car_cost = self.main_df[self.main_df['分类'].isin(["交通/私家车"])]
    
    car_repair_df = car_cost[car_cost['备注'].isin(["Bruce Auto repair","costco","Rear window blade"])]
    car_repair_df = car_repair_df.replace("交通/私家车","Car Repair")
    
    public_transit_df = self.main_df[self.main_df['备注'].isin(["Presto"])]
    
    gift_df = self.main_df[self.main_df["备注"].isin(["Mr.Surprise","Mr.Surprise","Panda Hobby","Shigum Hobbies","Claw & Kitty","Ebgames"])]
    gift_df = gift_df.replace("辦公用品","gift")
    
    office_supplies_df = self.main_df[self.main_df['分类'].isin(["辦公用品"])]
    filtered_office_supplies_df = office_supplies_df[~office_supplies_df["备注"].isin(["Bell Internet", "Mr.Surprise","Mr.Surprise","Panda Hobby","Shigum Hobbies","Claw & Kitty","Ebgames","Homesense","Lowes","Rona","Online courses","AWS Certified"])]
    
    training_df = self.main_df[self.main_df["备注"].isin(["Online courses","AWS Certified"])]
    training_df = training_df.replace("辦公用品", "Training")
    
    
    resturant_cost_df = self.main_df[self.main_df['分类'].isin(["餐饮"])]
    resturant_cost_df =resturant_cost_df[resturant_cost_df['金额']>20]
    
    
    final_result_df= pd.concat([
                          bank_account_fee_df,
                          travel_df,water_cost_df, hydro_cost_df,gas_cost_df,
                          telephone_line_cost_df,internet_df, 
                          parking_df,fuel_df, car_repair_df,public_transit_df,
                          gift_df, training_df,filtered_office_supplies_df,
                          resturant_cost_df
                          ])
    self.expected_df = final_result_df
    
  def get_grand_total_by_categories(self):
    # it is neccessary to reset the index or it should be done when genereating files
    self.grand_total_df = pd.DataFrame(self.expected_df.groupby(["分类"])['金额'].sum().reset_index())
    sum_of_total = {"分类":"Total","金额": sum(self.grand_total_df["金额"])}
    self.grand_total_df.loc[len(self.grand_total_df)] = sum_of_total

    
    # # self.grand_total_df = self.grand_total_df.append(sum_of_total,ignore_index=True)

    
    # # rows_to_be_added = pd.DataFrame(self.grand_total_df)
    # self.expected_df= pd.concat([self.expected_df, pd.DataFrame(sum_of_total,index=sum_of_total['分类']).reset_index()])
    print(self.grand_total_df)

    
  def export_to_excel(self, file_name):
    # self.expected_df.to_excel(file_name)
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
      # Write the original DataFrame to the first sheet
      self.expected_df.to_excel(writer, sheet_name='Original_Data', index=False)

      # Write the grand total DataFrame to a new sheet
      self.grand_total_df.to_excel(writer, sheet_name='Grand_Total', index=False)


if __name__ == '__main__':
  file_path = "new.xls"
  start_date = '2023-01-01'
  end_date = '2023-4-28'
  temp_obj = ReportProcessor(file_path,start_date,end_date )
  temp_obj.read_and_clean_data()
  temp_obj.get_desired_df()
  
  temp_obj.get_grand_total_by_categories()
  temp_obj.export_to_excel("temp_file.xlsx")
  # print(temp_obj.get_desired_df())
  
    
