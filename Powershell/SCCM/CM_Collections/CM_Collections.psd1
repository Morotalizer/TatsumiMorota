#
# Module manifest for module 'CM_Collections'
#
# Copyright="© ConsultingExperts. All rights reserved."
#

@{

ModuleVersion = '1.2'

GUID = '57eca2bd-b001-4278-975f-ada6fec0ca58 '

Author = 'Rob Looman'

CompanyName = 'ConsultingExperts'

Copyright = "© ConsultingExperts. All rights reserved."

Description = 'This module can import and export SCCM device collections'

PowerShellVersion = '3.0'

CLRVersion = '4.0'

NestedModules = @('CM_Collections.psm1')

RequiredModules = @('ConfigurationManager')

FunctionsToExport = @('Import-CMDeviceCollectionsFromXML',
                       'Export-CMDeviceCollectionsToXML',
                       'Import-CMUserCollectionsFromXML',
                       'Export-CMUserCollectionsToXML',
                       'Import-CMCollectionsFromXML',
                       'Export-CMCollectionsToXML')

# HelpInfoURI = ''

}

