# Office 365 DKIM Tool

This tool is for DKIM management in Microsoft Office 365

Tool to:
Create a CSV file for all CNAME's there has to be created to enable DKIM
Test DNS Settings
Enable DKIM

## Menu

1: Create DKIM CSV File

2: Test DKIM DNS

3: Enable DKIM

### Create DKIM CSV File

CSV file will show needed changes to your Public DNS

First part = domainname

Second part = Hostname to add

Third part = cname to match

example:

```csv
domain,hostname,cname
domain.com,selector1._domainkey,selector1-domain-com._domainkey.domain-com.onmicrosoft.com
domain.com,selector2._domainkey,selector2-domain-com._domainkey.domain-com.onmicrosoft.com
```