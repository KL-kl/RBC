from django import forms


class LoginForm(forms.Form):
    #这里所写的字段属性名一定要和前端form表单类元素当中的name属性值一致
    username = forms.CharField(required=True,min_length=4,error_messages={
        'required':'用户名必须填写',
        'min_length':'用户名长度最小是4'
    })
    password = forms.CharField(required=True, min_length=4, error_messages={
        'required': '密码必须填写',
        'min_length': '密码长度最小是4'
    })
