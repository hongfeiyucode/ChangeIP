# -*- coding: utf-8 -*-
import wmi
import random
print '欢迎使用红绯鱼の修改IP程序 <你现在位于宿舍楼>'
isjiaban=raw_input('你在加班室吗?(y/n)')
if isjiaban=='y' or isjiaban=='Y' :
    ip_3nd='161'
else :
    ip_3nd='160'
wmiService = wmi.WMI()
colNicConfigs = wmiService.Win32_NetworkAdapterConfiguration(IPEnabled = True)
if len(colNicConfigs) < 1:
    print '没有找到可用的网络适配器'
    exit()
for i in xrange(len(colNicConfigs)):
    if colNicConfigs[i].SettingID=='{5240CB7A-FA58-4A5A-8C1A-8C8300AAE22B}' :
        objNicConfig = colNicConfigs[i]
        break
print "以太网卡信息: "
#print objNicConfig.Index
#print objNicConfig.SettingID
#print objNicConfig.Description.encode("cp936")
print 'IP: ', ', '.join(objNicConfig.IPAddress)
if objNicConfig.DefaultIPGateway!=None :
    print '掩码: ', ', '.join(objNicConfig.IPSubnet)
    print '网关: ', ', '.join(objNicConfig.DefaultIPGateway)
    print 'DNS: ', ', '.join(objNicConfig.DNSServerSearchOrder)

def change(ip_last):
    print "随机选择IP为"+str(ip_last)
    arrIPAddresses = ['10.104.'+ip_3nd+'.'+str(ip_last)]
    arrSubnetMasks = ['255.255.255.0']
    arrDefaultGateways = ['10.104.'+ip_3nd+'.1']
    arrGatewayCostMetrics = [1]
    arrDNSServers = ['114.114.114.114', '8.8.8.8']
    intReboot = 0
    returnValue = objNicConfig.EnableStatic(IPAddress = arrIPAddresses, SubnetMask =arrSubnetMasks)
    if returnValue[0] == 0 or returnValue[0] == 1:
        print '设置IP成功'
        intReboot += returnValue[0]
    else:
        print '修改失败: IP或子网掩码设置发生错误'
        return
    returnValue = objNicConfig.SetGateways(DefaultIPGateway = arrDefaultGateways, GatewayCostMetric = arrGatewayCostMetrics)
    if returnValue[0] == 0 or returnValue[0] == 1:
        print '设置网关成功'
        intReboot += returnValue[0]
    else:
        print '修改失败: 网关设置发生错误'
        return

    returnValue = objNicConfig.SetDNSServerSearchOrder(DNSServerSearchOrder = arrDNSServers)
    if returnValue[0] == 0 or returnValue[0] == 1:
        print '设置DNS成功'
        intReboot += returnValue[0]
    else:
        print str(returnValue)+'修改失败: DNS设置发生错误'
        return

    # print 'IP: ', ', '.join(objNicConfig.IPAddress)
    # if objNicConfig.DefaultIPGateway!=None :
    #     print '掩码: ', ', '.join(objNicConfig.IPSubnet)
    #     print '网关: ', ', '.join(objNicConfig.DefaultIPGateway)
    #     print 'DNS: ', ', '.join(objNicConfig.DNSServerSearchOrder)

    if intReboot > 0:
        print '需要重新启动计算机'
    print '修改结束'


rand=random.randint(1, 254)
change(rand)
while(raw_input('是否可以上网?(y/n)')!='y'):
    rand=random.randint(1, 254)
    change(rand)