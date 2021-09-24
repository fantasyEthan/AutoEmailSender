# README

* `parse.py`为预处理文件，运行之后会生成加有班级信息的未打卡名单和`t4.html`文件

* `autoemail.py`为发送邮件程序，可以更改

  * `mail_sender`是自己的邮箱，建议是校园邮箱，如果更换其他邮箱需要相应更改`mail_host`
  * `mail_license`：如果使用校园邮箱，则直接填写邮箱密码即可

* 文件结构

  ```bash
  ├── 18级学生信息_邮箱.xlsx
  ├── README.md
  ├── autoemail.py
  ├── parse.py
  ├── t4.html
  ├── 未打卡名单.xls
  └── 班干信息.xlsx
  ```

  * 每天下载的新文件更名为`未打卡名单.xls`，放到目录下
  * 先运行`parse.py`，加有班级信息的未打卡名单发群里，`t4.html`文件打开看一下是不是当日未打卡的人，如果是，运行`autoemail.py`会自动发送邮件，等待即可

* 需要安装`pandas`
* 涉及敏感文件，如有需要可以联系s_lyu@smail.nju.edu.cn

