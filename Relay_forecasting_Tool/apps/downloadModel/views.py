import os
from django.utils.encoding import escape_uri_path
from django.http import HttpResponse
import xlrd
from xlutils.copy import copy

# Create your views here.
path = os.path.abspath('models')


# IP专网模板==================================================================
def downloadTest1(request):
    if request.method == 'POST':

        wb = xlrd.open_workbook(path + '\\IP网_metric.xls')
        wb2 = copy(wb)

        filename = request.POST.get('filename')
        if not filename:

            response = HttpResponse(content_type='application/octet-stream')
            response['Content-Disposition'] = 'attachment; filename="{0}"'.format(escape_uri_path('IP网_metric.xls'))
            wb2.save(response)
            return response
        else:

            response = HttpResponse(content_type='application/vnd.ms-excel')
            response['Content-Disposition'] = 'attachment; filename=' +filename + '.xls'
            wb2.save(response)
            return response

def downloadTest2(request):
    if request.method == 'POST':
        filename = request.POST.get('filename')
        wb = xlrd.open_workbook(path + '\\IP网_传输距离.xls')
        wb2 = copy(wb)
        if not filename:

            response = HttpResponse(content_type='application/octet-stream')
            response['Content-Disposition'] = 'attachment; filename="{0}"'.format(escape_uri_path('IP网_传输距离.xls'))
            wb2.save(response)
            return response
        else:


            response = HttpResponse(content_type='application/vnd.ms-excel')
            response['Content-Disposition'] = 'attachment; filename=' + filename + '.xls'
            wb2.save(response)
            return response


# CMNET网模板========================================================================
def downloadTest3(request):
    if request.method == 'POST':
        filename = request.POST.get('filename')

        wb = xlrd.open_workbook(path + '\\CMNET网_metric.xls')
        wb2 = copy(wb)
        if not filename:

            response = HttpResponse(content_type='application/octet-stream')
            response['Content-Disposition'] = 'attachment; filename="{0}"'.format(escape_uri_path('CMNET网_metric.xls'))
            wb2.save(response)
            return response
        else:

            response = HttpResponse(content_type='application/vnd.ms-excel')
            response['Content-Disposition'] = 'attachment; filename=' + filename + '.xls'
            wb2.save(response)
            return response


def downloadTest4(request):
    if request.method == 'POST':
        filename = request.POST.get('filename')

        wb = xlrd.open_workbook(path + '\\CMNET网_传输距离.xls')
        wb2 = copy(wb)

        if not filename:

            response = HttpResponse(content_type='application/octet-stream')
            response['Content-Disposition'] = 'attachment; filename="{0}"'.format(escape_uri_path('CMNET网_传输距离.xls'))
            wb2.save(response)
            return response
        else:

            response = HttpResponse(content_type='application/vnd.ms-excel')
            response['Content-Disposition'] = 'attachment; filename=' + filename + '.xls'
            wb2.save(response)
            return response
