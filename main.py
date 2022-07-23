
# lib for data processsing
import pandas as pd

# operating system library for get current path of program
import os



#FUNCTION - check format file when read
def check_file_format(file_path):
    # read file excel
    excel_data = pd.read_excel(file_path, index_col=None, header=None)
    info = []

    # put try-catch when format is wrong then show [file-name].xlsx WRONG
    try:
        info.append(excel_data[3])
    except:
        print(file_path + "wrong format")
        return 0
    return 1

#FUNCTION - read data from file excel
def get_data_from_file(file_path):

    # read file excel
    excel_data = pd.read_excel(file_path, index_col=None, header=None)

    # DATA IS AT COLUMN 4 AND ROW FROM 1,2,4- 37
    info = []

    # NAME
    info.append(excel_data[3][1])
    # POSITION
    info.append(excel_data[3][2])

    # ANOTHER DATA
    for i in range(4,38):
        info.append(excel_data[3][i])
    # get file xlsx name
    info.append(file_path.split("/data/")[1])
    return info

# get 'folder data' at the same directory of main.py
folder_data = os.getcwd() + "/data/" 

list_files = []

# get list file in folder "data"
for path in os.listdir(folder_data):

    # check if current path is a file
    if (os.path.isfile(os.path.join(folder_data, path))) and (path.find(".xlsx") != -1) and (path.find("~") == -1) :
        list_files.append(path)


data = []

# set header for tonghop.xlsx
data.append(["", "", 1 ,"", "", "", 2, "", "", "", 3, "", "", 4, "", "", "", "", 5, "", "", "", "", 6, "", "", "", 7, "", "", "", 8, "", "", "", ""])
data.append(["", "", 'Bằng Đại học (Cử nhân)',"", "", "", "Bằng Thạc sĩ (nếu có)", "", "", "", "Bằng Tốt nghiệp THPT", "", "", "Chứng chỉ ngoại ngữ", "", "", "", "", "Chứng chỉ tin học", "", "", "", "", "Chứng chỉ bồi dưỡng ", "kiến thức quản lý ", "nhà nước", "", "Chứng chỉ bồi dưỡng ", "tiêu chuẩn chức danh ", "nghề nghiệp viên chức", "", "Văn bằng hoặc ", "chứng chỉ khác ", "", "", ""])
data.append(["Họ và tên","Chức vụ","Chuyên ngành","Số hiệu","Nơi cấp","Ngày cấp","Chuyên ngành","Số hiệu","Nơi cấp","Ngày cấp","Số hiệu","Nơi cấp","Ngày cấp","Tên chứng chỉ","Trình độ","Số hiệu","Nơi cấp","Ngày cấp","Tên chứng chỉ","Trình độ","Số hiệu","Nơi cấp","Ngày cấp","Tên gọi","Số hiệu","Nơi cấp","Ngày cấp","Tên gọi","Số hiệu","Nơi cấp","Ngày cấp","Tên văn bằng","Trình độ","Số hiệu","Nơi cấp","Ngày cấp"])

# read data from excel file then add to a LIST
for file in list_files:
    # check format then add data read from excel file
    if(check_file_format(folder_data+file)):
        data.append(get_data_from_file(folder_data+file)) 
   
# transpose Datafram for writing to excel file
df = pd.DataFrame(data).transpose()



# writing to excel file
writer = pd.ExcelWriter('tonghop.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1',index=False,header=False)
writer.save()





# # Load the xlsx file
# excel_data = pd.read_excel('sales.xlsx')
# # Read the values of the file in the dataframe
# data = pd.DataFrame(excel_data, columns=['Sales Date', 'Sales Person', 'Amount'])
# # Print the content
# print("The content of the file is:\n", data)