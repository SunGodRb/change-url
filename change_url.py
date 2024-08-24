# %%
import requests
import pandas as pd
start_msg = input(">>> 请输入start后回车启动程序，否则输入其他任意内容后回车退出程序")
if start_msg == 'start':
    pass
else:
    exit()

# %%
def open_file(file_name):
    print('\n >>> 正在打开需要转换的短链文件')
    input_df = pd.read_excel(file_name)
    input_column = input_df.iloc[:,0]
    return input_column

try:
    input_column = open_file('change_url.xlsx')    
    result_column = []
    short_url_col = list(input_column)
    long_url_col = []
    total_nums = len(short_url_col)
    print(' >>> 短链列表共有{}个变量'.format(total_nums))

    print('\n >>> 将短链列表中的变量先转换为字符串类型，非字符串的类型将统一替换为error')
    converted_list = []
    for item in short_url_col:
        try:
            # 尝试将元素转换为字符串类型
            converted_item = str(item)
        except Exception as e:
            # 如果转换失败，则将元素替换为"错误参数"
            converted_item = "error"
        # 将转换后的元素添加到新的列表中
        converted_list.append(converted_item)


    recent_num = 0
    for url in converted_list:
        recent_num += 1
        print('\n >>> 转换进度：{}/{}'.format(recent_num,total_nums))
        print('    当前短链: {}'.format(url))
        if url.startswith('http'): 
            try:
                res = requests.head(url = url)
                res_res = res.headers.get('location')
                long_url_col.append(res_res)
                print("    当前长链: {}".format(res_res))
            except:
                print("    无法打开该链接")
                long_url_col.append("无法打开该链接")
        else:
            print("    参数错误")
            long_url_col.append('参数错误')
        

    output_col = ['短链','长链']
    output_df = pd.DataFrame(columns=output_col)
    output_df['短链'] = short_url_col
    output_df['长链'] = long_url_col
    print(output_df)
    output_df.to_excel('result_url.xlsx',index=False)
    input('\n >>> 输入Enter以退出')
except Exception as e:
    print(f"An error occurred: {e}")
    input("\n >>> 输入Enter以退出")


