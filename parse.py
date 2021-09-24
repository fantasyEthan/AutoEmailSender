import pandas as pd
import os
import datetime


def getEmail(file_path, name_path):
    dfile = pd.read_excel(file_path)
    dname = pd.read_excel(name_path)
    dfile = dfile[['学号', '姓名', '填报日期', '填报状态', '是否两天未打卡']]
    dfile = dfile.loc[lambda df:df['填报状态'] == '未填报']
    for fidx, frows in dfile.iterrows():
        for nidx, nrows in dname.iterrows():
            if (frows['学号'] == nrows['学号']):
                dfile.loc[fidx, '邮箱'] = dname.loc[nidx, '常用电子邮箱']
                dfile.loc[fidx, '班级'] = dname.loc[nidx, '班级']
                dfile.loc[fidx, 'Tel'] = dname.loc[nidx, '手机号码']
    dfile['Tel'] = dfile['Tel'].apply(lambda _: str(_))
    return dfile


def classsifyClass(dfile):
    df1 = dfile.loc[df['班级'] == "一班"]
    df2 = dfile.loc[df['班级'] == "二班"]
    df3 = dfile.loc[df['班级'] == "三班"]
    df4 = dfile.loc[df['班级'] == "四班"]
    df5 = dfile.loc[df['班级'] == "五班"]
    return df1, df2, df3, df4, df5


def get_df_html(df):
    df_html = df.to_html(escape=False)
    head = \
        """
        <head>
            <meta charset="utf-8">
            <STYLE TYPE="text/css" MEDIA=screen>

                table.dataframe {
                    border-collapse: collapse;
                    border: 2px solid #a19da2;
                    /*居中显示整个表格*/
                    margin: auto;
                }

                table.dataframe thead {
                    border: 2px solid #91c6e1;
                    background: #f1f1f1;
                    padding: 10px 10px 10px 10px;
                    color: #333333;
                }

                table.dataframe tbody {
                    border: 2px solid #91c6e1;
                    padding: 10px 10px 10px 10px;
                }

                table.dataframe tr {

                }

                table.dataframe th {
                    vertical-align: top;
                    font-size: 14px;
                    padding: 10px 10px 10px 10px;
                    color: #105de3;
                    font-family: arial;
                    text-align: center;
                }

                table.dataframe td {
                    text-align: center;
                    padding: 10px 10px 10px 10px;
                }

                body {
                    font-family: 宋体;
                }

                h1 {
                    color: #5db446
                }

                div.header h2 {
                    color: #0002e3;
                    font-family: 黑体;
                }

                div.content h2 {
                    text-align: center;
                    font-size: 28px;
                    text-shadow: 2px 2px 1px #de4040;
                    color: #fff;
                    font-weight: bold;
                    background-color: #8A2BE2;
                    line-height: 1.5;
                    margin: 20px 0;
                    box-shadow: 10px 10px 5px #888888;
                    border-radius: 5px;
                }

                h3 {
                    font-size: 22px;
                    background-color: rgba(0, 2, 227, 0.71);
                    text-shadow: 2px 2px 1px #de4040;
                    color: rgba(239, 241, 234, 0.99);
                    line-height: 1.5;
                }

                h4 {
                    color: #e10092;
                    font-family: 楷体;
                    font-size: 20px;
                    text-align: center;
                }

                td img {
                    /*width: 60px;*/
                    max-width: 300px;
                    max-height: 300px;
                }

            </STYLE>
        </head>
        """
    # 构造模板的附件
    body = \
        """
        <body>
        
        <div align="center" class="header">
            <!--标题部分的信息-->
            <!--<h1 align="center">我的python邮件，使用了Dataframe转为table </h1> -->
        </div>

        <hr> 

        <div class="content">
            <!--正文内容-->
            <h4>各位班干，今日未在南大APP上打卡名单如下，请尽快联系本班同学进行填报，负责点名的同学记得10点在群里进行回复，辛苦大家！</h4>
            <hr>
            <h2>今日尚未打卡名单</h2>
            <div>
                <h4></h4>
                {df_html}
            </div>
        </body>
        """.format(df_html=df_html)
    html_msg = "<html>" + head + body + "</html>"
    fout = open('t4.html', 'w', encoding='UTF-8', newline='')
    fout.write(html_msg)
    return html_msg


if __name__ == '__main__':
    dir_path = os.getcwd()
    file_path = os.path.join(dir_path, '未打卡名单.xls')
    name_path = os.path.join(dir_path, '18级学生信息_邮箱.xlsx')
    df = getEmail(file_path, name_path)
    df1, df2, df3, df4, df5 = classsifyClass(df)
    time = datetime.datetime.today()
    time = time.strftime("%Y-%m-%d")
    df1, df2, df3, df4, df5 = classsifyClass(df)
    html_msg = get_df_html(df)
    # df1.to_excel(os.path.join(dir_path, '1班' + time + '未打卡名单.xlsx'))
    # df2.to_excel(os.path.join(dir_path, '2班' + time + '未打卡名单.xlsx'))
    # df3.to_excel(os.path.join(dir_path, '3班' + time + '未打卡名单.xlsx'))
    # df4.to_excel(os.path.join(dir_path, '4班' + time + '未打卡名单.xlsx'))
    # df5.to_excel(os.path.join(dir_path, '5班' + time + '未打卡名单.xlsx'))
    writer = pd.ExcelWriter(os.path.join(
        dir_path, '18级' + time + '未打卡名单.xlsx'), engine='xlsxwriter')
    df.to_excel(writer, startrow=0, startcol=0,
                sheet_name='Sheet1', index=False)
    ws = writer.sheets['Sheet1']

    for i, col in enumerate(df.columns):
        # find length of column i
        column_len = df[col].astype(str).str.len().max()
        # Setting the length if the column header is larger
        # than the max column value length
        print(col)
        print(len(col))
        column_len = max(column_len+5, len(col)+10)
        # set the column length
        ws.set_column(i, i, column_len)
    writer.save()
