from django.db import models
from datetime import datetime
from django.contrib.auth.models import User

class ProjectInfo(models.Model):
    id = models.IntegerField(primary_key=True)
    name = models.CharField(max_length=30,verbose_name='项目名称')
    pro_type = models.IntegerField(verbose_name='中继仿真模型')
    pm = models.CharField(max_length=10, verbose_name='平面')
    xs = models.IntegerField(default=80,verbose_name='业务复用系数')
    nodefile = models.CharField(max_length=200, verbose_name='节点信息表')
    businessfile = models.CharField(max_length=500, verbose_name='业务表')
    PtoPflowfile = models.CharField(max_length=200, verbose_name='点到点流量表')
    metricfile = models.CharField(max_length=200, verbose_name='Metric表')
    distancefile = models.CharField(max_length=200, verbose_name='传输距离表')
    Relaydirectionfile = models.CharField(max_length=200, verbose_name='中继方向表')
    TEfile = models.CharField(max_length=200, verbose_name='TE主备路由表')
    Relaystatusfile = models.CharField(max_length=200, verbose_name='中继现状表')
    createBy = models.ForeignKey(User, verbose_name='操作用户')
    add_time = models.DateTimeField(default=datetime.now, verbose_name='添加时间')

    def __str__(self):
        return self.name

    class Meta:
        db_table = 'e_ProjectInfo'
        verbose_name = '项目信息'
        verbose_name_plural = verbose_name


class CNodes_detail(models.Model):
    '''节点信息'''
    province = models.CharField(max_length=20, verbose_name='省份')
    city = models.CharField(max_length=20, verbose_name='城市')
    office_address1 = models.CharField(max_length=50, verbose_name='现状局址')
    building_no1 = models.CharField(max_length=20, verbose_name='现状楼号')
    floor1 = models.CharField(max_length=20, verbose_name='现状楼层')
    room_num1 = models.CharField(max_length=20, verbose_name='现状机房编号')
    plane1 = models.CharField(max_length=10, verbose_name='现状流量平面')
    network_level1 = models.CharField(max_length=20, verbose_name='现状网络层次')
    part1 = models.CharField(max_length=20, verbose_name='现状角色')
    part_subdivide1 = models.CharField(max_length=20, verbose_name='现状角色-细分')
    devicename1 = models.CharField(max_length=30, verbose_name='现状设备名称')
    device_comp1 = models.CharField(max_length=30, verbose_name='现状设备厂商')
    unit_type1 = models.CharField(max_length=30, verbose_name='现状设备型号')
    device_state1 = models.CharField(max_length=30, verbose_name='现状设备厂商')
    program = models.CharField(max_length=30, verbose_name='建设方案')
    office_address2 = models.CharField(max_length=50, verbose_name='局址')
    building_no2 = models.CharField(max_length=20, verbose_name='楼号')
    floor2 = models.CharField(max_length=20, verbose_name='楼层')
    room_num2 = models.CharField(max_length=20, verbose_name='机房编号')
    plane2 = models.CharField(max_length=10, verbose_name='流量平面')
    network_level2 = models.CharField(max_length=20, verbose_name='网络层次')
    part2 = models.CharField(max_length=20, verbose_name='角色')
    part_subdivide2 = models.CharField(max_length=20, verbose_name='角色-细分')
    devicename2 = models.CharField(max_length=30, verbose_name='设备名称')
    device_comp2 = models.CharField(max_length=30, verbose_name='设备厂商')
    unit_type2 = models.CharField(max_length=30, verbose_name='设备型号')
    device_state2 = models.CharField(max_length=30, verbose_name='设备厂商')

    # admin显示节点的名称
    def __str__(self):
        return self.devicename1,self.devicename2

    class Meta:
        db_table = 'e_CNodes_detail'
        verbose_name = 'CMNET节点信息'
        verbose_name_plural = 'CMNET节点信息'


class CMetricTable(models.Model):

    '''metric'''
    mid = models.BigIntegerField(verbose_name='metric操作ID')
    fid = models.BigIntegerField(verbose_name='metric文件ID')
    file_path = models.CharField(max_length=200, verbose_name='metric文件路径')
    filename = models.CharField(max_length=100, verbose_name='metric文件名称')
    proid = models.ForeignKey(ProjectInfo, verbose_name='所属项目')
    createdBy = models.ForeignKey(User, verbose_name='操作用户')
    create_time = models.DateTimeField(auto_now_add=True, verbose_name='创建时间')


    class Meta:
        db_table = 'e_CMetricTable'
        verbose_name = 'CMNET_metric操作'
        verbose_name_plural = 'CMNET_metric操作'


class CPtoPoint_flow(models.Model):

    '''点到点流量'''
    flowid = models.BigIntegerField(verbose_name='点到点流量操作ID')
    fid = models.BigIntegerField(verbose_name='点到点流量文件ID')
    file_path = models.CharField(max_length=200, verbose_name='点到点流量文件路径')
    filename = models.CharField(max_length=100, verbose_name='点到点流量文件名称')
    proid = models.ForeignKey(ProjectInfo, verbose_name='所属项目')
    createdBy = models.ForeignKey(User, verbose_name='操作用户')
    create_time = models.DateTimeField(auto_now_add=True, verbose_name='创建时间')

    class Meta:
        db_table = 'e_CPtoPoint_flow'
        verbose_name = 'CMNET_点到点流量'
        verbose_name_plural = 'CMNET_点到点流量'


class CPtoPoint_route(models.Model):

    '''点点路由'''
    routeid = models.BigIntegerField(verbose_name='点点路由操作ID')
    fid = models.BigIntegerField(verbose_name='点点路由文件ID')
    file_path = models.CharField(max_length=200, verbose_name='点点路由文件路径')
    filename = models.CharField(max_length=100, verbose_name='点点路由文件名称')
    proid = models.ForeignKey(ProjectInfo, verbose_name='所属项目')
    createdBy = models.ForeignKey(User, verbose_name='操作用户')
    create_time = models.DateTimeField(auto_now_add=True, verbose_name='创建时间')

    class Meta:
        db_table = 'e_CPtoPoint_route'
        verbose_name = 'CMNET_点点路由'
        verbose_name_plural = 'CMNET_点点路由'


class CMalfunction(models.Model):
    '''CMNET网故障'''
    raleyCircuitid = models.BigIntegerField(verbose_name='故障操作ID')
    fid = models.BigIntegerField(verbose_name='故障文件ID')
    file_path = models.CharField(max_length=200, verbose_name='故障文件路径')
    filename = models.CharField(max_length=100, verbose_name='故障文件名称')
    proid = models.ForeignKey(ProjectInfo, verbose_name='所属项目')
    createdBy = models.ForeignKey(User, verbose_name='操作用户')
    create_time = models.DateTimeField(auto_now_add=True, verbose_name='创建时间')

    class Meta:
        db_table = 'e_CMalfunction'
        verbose_name = 'CMNET_故障轮巡'
        verbose_name_plural = 'CMNET_故障轮巡'

class CRelay_traffic(models.Model):

    '''中继转发流量'''
    raleyid = models.BigIntegerField(verbose_name='中继转发流量操作ID')
    fid = models.BigIntegerField(verbose_name='中继转发流量文件ID')
    file_path = models.CharField(max_length=200, verbose_name='中继转发流量文件路径')
    filename = models.CharField(max_length=100, verbose_name='中继转发流量文件名称')
    proid = models.ForeignKey(ProjectInfo, verbose_name='所属项目')
    createdBy = models.ForeignKey(User, verbose_name='操作用户')
    create_time = models.DateTimeField(auto_now_add=True, verbose_name='创建时间')

    class Meta:
        db_table = 'e_CRelaying_traffic'
        verbose_name = 'CMNET_中继转发流量'
        verbose_name_plural = 'CMNET_中继转发流量'


class CRelay_circuit(models.Model):

    '''中继电路'''
    raleyCircuitid = models.BigIntegerField(verbose_name='中继电路操作ID')
    fid = models.BigIntegerField(verbose_name='中继电路文件ID')
    file_path = models.CharField(max_length=200, verbose_name='中继电路文件路径')
    filename = models.CharField(max_length=100, verbose_name='中继电路文件名称')
    proid = models.ForeignKey(ProjectInfo, verbose_name='所属项目')
    createdBy = models.ForeignKey(User, verbose_name='操作用户')
    create_time = models.DateTimeField(auto_now_add=True, verbose_name='创建时间')

    class Meta:
        db_table = 'e_CRelay_circuit'
        verbose_name = 'CMNET_中继电路'
        verbose_name_plural = 'CMNET_中继电路'