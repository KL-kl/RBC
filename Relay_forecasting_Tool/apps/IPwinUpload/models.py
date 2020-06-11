from django.db import models
from CMNETwinUpload.models import ProjectInfo
from django.contrib.auth.models import User

# Create your models here.
class IpNodes_detail(models.Model):
    '''节点信息'''
    province1 = models.CharField(max_length=20, verbose_name='本工程前省份')
    node1 = models.CharField(max_length=20, verbose_name='本工程前节点')
    office_address1 = models.CharField(max_length=50, verbose_name='本工程前局址')
    building_no1 = models.CharField(max_length=20, verbose_name='本工程前楼号')
    floor1 = models.CharField(max_length=20, verbose_name='本工程前楼层')
    room_num1 = models.CharField(max_length=20, verbose_name='本期前机房编号')
    plane1 = models.CharField(max_length=10, verbose_name='本工程前流量平面')
    part1 = models.CharField(max_length=20, verbose_name='本工程前角色')
    network_level1 = models.CharField(max_length=20, verbose_name='本期前网络层次')
    province2 = models.CharField(max_length=20, verbose_name='本工程后省份')
    node2 = models.CharField(max_length=20, verbose_name='本工程后节点')
    office_address2 = models.CharField(max_length=50, verbose_name='本工程后局址')
    building_no2 = models.CharField(max_length=20, verbose_name='本工程后楼号')
    floor2 = models.CharField(max_length=20, verbose_name='本工程后楼层')
    room_num2 = models.CharField(max_length=20, verbose_name='本期后机房编号')
    plane2 = models.CharField(max_length=10, verbose_name='本工程后流量平面')
    part2 = models.CharField(max_length=20, verbose_name='本工程后角色')
    network_level2 = models.CharField(max_length=20, verbose_name='本期后网络层次')
    program = models.CharField(max_length=30, verbose_name='建设方案')

    # admin显示节点的名称
    def __str__(self):
        return self.node1,self.node2

    class Meta:
        db_table = 'e_IpNodes_detail'
        verbose_name = 'IP节点信息'
        verbose_name_plural = 'IP节点信息'


class IPMetricTable(models.Model):

    '''metric'''
    mid = models.BigIntegerField(verbose_name='metric操作ID')
    fid = models.BigIntegerField(verbose_name='metric文件ID')
    file_path = models.CharField(max_length=200, verbose_name='metric文件路径')
    filename = models.CharField(max_length=100, verbose_name='metric文件名称')
    proid = models.ForeignKey(ProjectInfo, verbose_name='所属项目')
    createdBy = models.ForeignKey(User, verbose_name='操作用户')
    create_time = models.DateTimeField(auto_now_add=True, verbose_name='创建时间')


    class Meta:
        db_table = 'e_IPMetricTable'
        verbose_name = 'IP_metric操作'
        verbose_name_plural = 'IP_metric操作'


class IPPtoPoint_flow(models.Model):

    '''点到点流量'''
    flowid = models.BigIntegerField(verbose_name='点到点流量操作ID')
    fid = models.BigIntegerField(verbose_name='点到点流量文件ID')
    file_path = models.CharField(max_length=200, verbose_name='点到点流量文件路径')
    filename = models.CharField(max_length=100, verbose_name='点到点流量文件名称')
    proid = models.ForeignKey(ProjectInfo, verbose_name='所属项目')
    createdBy = models.ForeignKey(User, verbose_name='操作用户')
    create_time = models.DateTimeField(auto_now_add=True, verbose_name='创建时间')

    class Meta:
        db_table = 'e_IPPtoPoint_flow'
        verbose_name = 'IP_点到点流量'
        verbose_name_plural = 'IP_点到点流量'


class IPPtoPoint_route(models.Model):

    '''点点路由'''
    routeid = models.BigIntegerField(verbose_name='点点路由操作ID')
    fid = models.BigIntegerField(verbose_name='点点路由文件ID')
    file_path = models.CharField(max_length=200, verbose_name='点点路由文件路径')
    filename = models.CharField(max_length=100, verbose_name='点点路由文件名称')
    proid = models.ForeignKey(ProjectInfo, verbose_name='所属项目')
    createdBy = models.ForeignKey(User, verbose_name='操作用户')
    create_time = models.DateTimeField(auto_now_add=True, verbose_name='创建时间')

    class Meta:
        db_table = 'e_IPPtoPoint_route'
        verbose_name = 'IP_点点路由'
        verbose_name_plural = 'IP_点点路由'


class IPRelay_traffic(models.Model):

    '''中继转发流量'''
    raleyid = models.BigIntegerField(verbose_name='中继转发流量操作ID')
    fid = models.BigIntegerField(verbose_name='中继转发流量文件ID')
    file_path = models.CharField(max_length=200, verbose_name='中继转发流量文件路径')
    filename = models.CharField(max_length=100, verbose_name='中继转发流量文件名称')
    proid = models.ForeignKey(ProjectInfo, verbose_name='所属项目')
    createdBy = models.ForeignKey(User, verbose_name='操作用户')
    create_time = models.DateTimeField(auto_now_add=True, verbose_name='创建时间')

    class Meta:
        db_table = 'e_IPRelaying_traffic'
        verbose_name = 'IP_中继转发流量'
        verbose_name_plural = 'IP_中继转发流量'


class IPRelay_circuit(models.Model):

    '''中继电路'''
    raleyCircuitid = models.BigIntegerField(verbose_name='中继电路操作ID')
    fid = models.BigIntegerField(verbose_name='中继电路文件ID')
    file_path = models.CharField(max_length=200, verbose_name='中继电路文件路径')
    filename = models.CharField(max_length=100, verbose_name='中继电路文件名称')
    proid = models.ForeignKey(ProjectInfo, verbose_name='所属项目')
    createdBy = models.ForeignKey(User, verbose_name='操作用户')
    create_time = models.DateTimeField(auto_now_add=True, verbose_name='创建时间')

    class Meta:
        db_table = 'e_IPRelay_circuit'
        verbose_name = 'IP_中继电路'
        verbose_name_plural = 'IP_中继电路'


class IPMalfunction(models.Model):
    '''IP网故障'''
    raleyCircuitid = models.BigIntegerField(verbose_name='中继电路操作ID')
    fid = models.BigIntegerField(verbose_name='中继电路文件ID')
    file_path = models.CharField(max_length=200, verbose_name='中继电路文件路径')
    filename = models.CharField(max_length=100, verbose_name='中继电路文件名称')
    proid = models.ForeignKey(ProjectInfo, verbose_name='所属项目')
    createdBy = models.ForeignKey(User, verbose_name='操作用户')
    create_time = models.DateTimeField(auto_now_add=True, verbose_name='创建时间')

    class Meta:
        db_table = 'e_IPMalfunction'
        verbose_name = 'IP_故障轮巡'
        verbose_name_plural = 'IP_故障轮巡'


class IP_TE(models.Model):
    '''TE主备'''
    main_route = models.CharField(max_length=100, verbose_name='主路由')
    standby_route= models.CharField(max_length=100, verbose_name='备路由')
    update_time = models.DateTimeField(auto_now=True, verbose_name='更新时间')

    class Meta:
        db_table = 'e_IP_TE'
        verbose_name = 'IP_TE主备LSP'
        verbose_name_plural = 'IP_TE主备LSP'

