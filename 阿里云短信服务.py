from aliyunsdkcore.client import AcsClient
from aliyunsdkcore.request import CommonRequest


"""
安装阿里云SDK:
pip3 install aliyun-python-sdk-core
pip3 install aliyun-python-sdk-dysmsapi

版本：2024_7_1     版本时间：2024.7.1

使用阿里云发送短信服务


"""

# 创建AcsClient实例，需替换为您自己的AccessKey信息
ACCESS_KEY_ID = '*******************'  #在阿里云控制台创建AccessKey时自动生成的一对访问密钥，上面保存的AccessKey
ACCESS_KEY_SECRET = '*******************'  #在阿里云控制台创建AccessKey时自动生成的一对访问密钥AccessKey
SIGN_NAME = '老王出品'  # 短信签名
template_code = 'SMS_468770144' #短信模板CODE
PhoneNumber = '13242281907'  # 绑定的测试手机号
acs_client = AcsClient(ACCESS_KEY_ID, ACCESS_KEY_SECRET, template_code)

# 创建CommonRequest实例
request = CommonRequest()

# 设置请求参数,下面这5行其实不用动
request.set_accept_format('json')  # 设置API响应格式的方法
request.set_domain('dysmsapi.aliyuncs.com')  # 设置API的域名的方法
request.set_method('POST')  # 设置API请求方法
request.set_version('2017-05-25')  # 设置API版本号
request.set_action_name('SendSms')  # 设置API操作名

# 设置短信模板参数
request.add_query_param('PhoneNumbers', PhoneNumber)
request.add_query_param('SignName', SIGN_NAME)
request.add_query_param('TemplateCode', template_code)
request.add_query_param('TemplateParam', '{"name":"江门电信白石机房1"}')

# 发送短信请求并获取返回结果
response = acs_client.do_action_with_exception(request)

print(response)
