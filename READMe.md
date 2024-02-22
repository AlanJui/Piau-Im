# 專案指引

本專案目標：為漢字建立河洛話標音工具，使部落格網誌便於分享含漢字標音之古文、詩詩文章。

漢字標音支援以下方式：
- 十五音 (以彙集雅俗通為標準)
- 方音符號 (TPS)
- 白話字教會羅馬拚音 (POJ)
- 台羅拚音 (TL)
- 閩南語拚音方案 (BP)

<!-- markdownlint-disable MD024 no-duplicate-heading/no-duplicate-header -->
<!-- markdownlint-disable MD043 required-headings/required-headers  -->

# 《專案指引》

## 專案摘要

建置「河洛話」工具網站，用以滿足以下應用需求：

- 查詢漢字的彙集雅俗通十五音標音
- 查詢反切的台羅音標
- 收集漢字在河洛話之標準用字

## 專案目錄結構

- Django Templates ： 《Project》/Templates/
- 靜態檔案： 《Project》/static/

## 建置專案軟體開發環境作業流程

### (1) 自 GitHub 下載專案原始碼

```sh
$ cd ~/workspace/Ho-Lok-Ue/
$ git clone git@github.com:AlanJui/ho-lok-ue-tools.git
$ cd ho-lok-ue-tools
$ ls -al
total 16
drwxr-xr-x   8 alanjui  staff   256  2 19 14:23 .
drwxr-xr-x  42 alanjui  staff  1344  2 19 14:19 ..
-rw-r--r--   1 alanjui  staff     7  2 19 14:22 .python-version
drwxr-xr-x   6 alanjui  staff   192  2 19 14:23 .venv
-rw-r--r--   1 alanjui  staff     0  2 19 14:19 README.md
drwxr-xr-x   3 alanjui  staff    96  2 19 14:19 ex_2023_02
-rw-r--r--   1 alanjui  staff   308  2 19 14:19 pyproject.toml
drwxr-xr-x   3 alanjui  staff    96  2 19 14:19 tests
```

### (2) 建置專案使用之 Python

```sh
$ cat .python-version
3.10.6
$ pyenv install 3.10.6  # 安裝專案所需使用 Python 直譯器
$ pyenv versions

$ pyenv local 3.10.6
$ pyenv version
$ python --version
Python 3.10.6
```

### (3) 啟用專案虛擬環境並安裝所需套件

```sh
$ python -m venv .venv
$ ls -al .venv

$ poetry shell
$ poetry install
$ ls .venv/lib/python3.10/site-packages/
```

### (4) 確認專案開發環境已完成設定作業

```sh
❯ python manage.py runserver
Watching for file changes with StatReloader
Performing system checks...

System check identified no issues (0 silenced).
February 19, 2023 - 06:39:15
Django version 4.1.7, using settings 'web_app.settings'
Starting development server at http://127.0.0.1:8000/
Quit the server with CONTROL-C.
```

## 參考

### 引用 Bootstrap

```sh
yarn add bootstrap
```

```sh
yarn global add bootstrap
la ~/.config/yarn/global/node_modules/bootstrap/
```

### Django 參考

- 程式碼來
  源：[《跟著老齊學 Python: Django 實戰》](https://github.com/qiwsir/DjangoPracticeProject)

- [Django V3.2 文档内容](https://docs.djangoproject.com/zh-hans/3.2/contents/)

### Django 使用者管理

#### **【官網文件】**

- [Django 中的用户认证¶](https://docs.djangoproject.com/zh-hans/3.2/topics/auth/)

- [使用 Django 的验证系统](https://docs.djangoproject.com/zh-hans/3.2/topics/auth/default/)

#### **【分享文件】**

- [Django/用戶認證](https://zh.wikibooks.org/zh-tw/Django/%E7%94%A8%E6%88%B7%E8%AE%A4%E8%AF%81)

- [Part 3: Django Reset Password and Change Password](https://techpluslifestyle.com/technology/django-reset-password/)

- [How to change a user’s password in Django](https://canbaysal.com/2018/10/how-to-change-a-users-password-in-django/)