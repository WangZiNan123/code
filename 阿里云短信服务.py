# -*- coding: utf-8 -*-
import os
import sys
from typing import List

from alibabacloud_dysmsapi20170525.client import Client as Dysmsapi20170525Client
from alibabacloud_tea_openapi import models as open_api_models
from alibabacloud_dysmsapi20170525 import models as dysmsapi_20170525_models
from alibabacloud_tea_util import models as util_models
from alibabacloud_tea_util.client import Client as UtilClient


class Sample:
    def __init__(self):
        pass

    @staticmethod
    def create_client() -> Dysmsapi20170525Client:
        config = open_api_models.Config(
            access_key_id=os.environ.get('ALIBABA_CLOUD_ACCESS_KEY_ID'),
            access_key_secret=os.environ.get('ALIBABA_CLOUD_ACCESS_KEY_SECRET')
        )
        config.endpoint = 'dysmsapi.aliyuncs.com'
        return Dysmsapi20170525Client(config)

    @staticmethod
    def send_sms_to_device(device_names: str) -> None:
        client = Sample.create_client()
        send_sms_request = dysmsapi_20170525_models.SendSmsRequest(
            sign_name='老王出品',  # 替换为您的签名名称
            template_code='SMS_468770144',  # 替换为您的模板代码
            phone_numbers='13242281907',  # 替换为接收短信的手机号码
            template_param=f'{{"name":"{device_names}"}}'  # 设备名称作为模板参数
        )
        runtime = util_models.RuntimeOptions()
        try:
            # 尝试发送短信的代码
            client.send_sms_with_options(send_sms_request, runtime)
        except Exception as error:
            # 打印错误的描述信息
            print(str(error))

    @staticmethod
    def main(device_names: List[str]) -> None:
        for device_name in device_names:
            Sample.send_sms_to_device(device_name)


if __name__ == '__main__':
   
    device_names = ['台山川岛长堤']
    Sample.main(device_names)
