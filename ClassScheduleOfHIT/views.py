#!/usr/bin/env python
# encoding:utf-8

from django.http import HttpResponse
from django.shortcuts import render
from django import forms

from utils.hit import sso_login
from utils.hit import download_schedule
from utils.convert import parseExcel

import random
import datetime
import string


class UploadFileForm(forms.Form):
    file = forms.FileField()
    semester_start_date = forms.DateField(help_text="(2018-02-26)")


def random_string(length):
    return ''.join(random.choice(string.letters) for _ in range(length))


def store(request):
    if request.method == "POST":
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            file = request.FILES['file']
            filename = "ClassScheduleOfHIT/uploads/%s.xls" % (random_string(0x10))
            with open(filename, "wb+") as f:
                f.write(file.read())
            print "Stored! %s" % filename
            parseExcel(filename, "ClassScheduleOfHIT/outputfiles/%s.cvs" % (random_string(0x10)),
                       datetime.datetime.strptime(request.POST.get("semester_start_date"), "%Y-%m-%d"))
            return HttpResponse("Import finished!")
        else:
            return HttpResponse("Form error!")

    else:
        form = UploadFileForm()
    return render(
        request,
        'upload_form.html',
        {
            'form': form,
            'title': 'Excel file uploads and download example',
            'header': ('Please choose any excel file ' +
                       'from your cloned repository:')
        })


class LoginForm(forms.Form):
    username = forms.CharField(max_length=0x10)
    password = forms.CharField(widget=forms.PasswordInput, max_length=0x20)
    semester_year = forms.IntegerField(min_value=2014)
    semester = forms.IntegerField()


def login(request):
    if request.method == "POST":
        form = LoginForm(request.POST)
        if form.is_valid():
            result = ""
            username = request.POST.get('username')
            password = request.POST.get('password')
            semester_year = request.POST.get('POST.semester_year')
            semester = request.POST.get('semester')
            result += "正在登录..." + "</br>"
            try:
                login_result = sso_login(username, password)
                if login_result[0]:
                    result += "登录成功..." + "</br>"
                    result += "正在下载课表..." + "</br>"
                    filename = "ClassScheduleOfHIT/uploads/%s.xls" % (random_string(0x10))
                    download_schedule(semester_year, semester, filename)
                    result += "下载成功..." + "</br>"
                else:
                    result += "登录失败...(%s)" % ((login_result[1]).encode("utf-8")) + "</br>"
            except Exception as e:
                print e
            return HttpResponse(result)
    else:
        form = LoginForm(initial={"semester_year": 2017, "semester": 2})
    return render(
        request,
        'upload_form.html',
        {
            'form': form,
            'title': 'Excel file uploads and download example',
            'header': ('Please choose any excel file ' +
                       'from your cloned repository:')
        }
    )
