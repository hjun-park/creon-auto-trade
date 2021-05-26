import win32com.client

instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
instMarketEye = win32com.client.Dispatch("CpSysDib.MarketEye")

tarketCodeList = instCpCodeMgr.GetGroupCodeList(5)

# Get PER
instMarketEye.SetInputValue(0, 67)
instMarketEye.SetInputValue(1, tarketCodeList)

# BlockRequest
instMarketEye.BlockRequest()

# GetHeaderValue
numStock = instMarketEye.GetHeaderValue(2)


# 업종별 코드
industryCodeList = instCpCodeMgr.GetIndustryList()

for industryCode in industryCodeList:
    print(industryCode, instCpCodeMgr.GetIndustryName(industryCode))


# 업종
instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
tarketCodeList = instCpCodeMgr.GetGroupCodeList(5)

for code in tarketCodeList:
    print(code, instCpCodeMgr.CodeToName(code))


# GetData
sumPer = 0
for i in range(numStock):
    sumPer += instMarketEye.GetDataValue(0, i)

print("Average PER: ", sumPer / numStock)

