#!/usr/bin/env python
# -*- coding: utf-8 -*-

import requests
import os
import config

line_notify_token = config.LINE_NOTIFY_TOKEN

def send_line_notify(notification_message):
    line_notify_api = 'https://notify-api.line.me/api/notify'
    headers = {'Authorization': f'Bearer {line_notify_token}'}
    data = {'message': f'message: {notification_message}'}
    requests.post(line_notify_api, headers = headers, data = data)
