import pandas as pd
import datetime


class Logger:
    def __init__(self, log_file, columns):
        self.log_file = log_file
        self.columns = columns
        self.log_df = pd.DataFrame(columns=columns)

    def log(self, data):
        if len(data) != len(self.columns) - 1:
            raise ValueError("Data length does not match the number of columns.")
        timestamp = datetime.datetime.now()
        log_data = {'Update': timestamp, **data}
        # pandas DataFrame中添加一行新数据,self.log_df 是一个pandas DataFrame对象，可能用于存储日志数据,len(self.log_df) 获取当前DataFrame的行数
        # self.log_df.loc[len(self.log_df)] 指定了一个新行，其索引等于当前DataFrame的长度
        self.log_df.loc[len(self.log_df)] = log_data

    def save_log_to_excel(self):
        self.log_df.to_excel(self.log_file, index=False, merge_cells=False)

# # 创建Logger对象，传递列名作为参数
# logger = Logger("log.csv", ["Timestamp", "Message", "Value"])
#
# # 记录日志
# logger.log(["This is a log message", 42])
# logger.log(["Another log message", 123])
#
# # 保存日志到CSV文件
# logger.save_log_to_csv()
