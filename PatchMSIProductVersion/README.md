#

PatchMSIProductVersion.vbs

There is an issue regarding a version mismatch between the Outlook Web Access (OWA) S/MIME Control in Exchange 2007 SP3. This script will change the version number and resolve the issue.

## Description

In the recently released Exchange 2007 Service Pack 3, there's a version mismatch between the Outlook Web Access (OWA) S/MIME Control, an Active X control used to provide S/MIME support in OWA. After you install SP3, users who have the control installed will get prompted to install the latest version of the control. The way this works – the code compares the "Version" property of the client S/MIME control (MIMECTL.DLL) on the user's computer with the ProductVersion property of the MSI file (OWASMIME.MSI) on the Client Access Server. During the released SP3 build, the version of the MSI file was incremented to 8.3.83.2. However, due to an error, the DLL file in the MSI retained its old version number (8.3.83.0). As a result, when Outlook Web Access users using Internet Explorer use S/MIME functionality, they get the same prompt to upgrade the S/MIME control even after they've upgraded. One way to resolve this issue is to download and run the PatchMSIProductVersion.vbs script which changes the version number. 
For more information on this download, see: http://msexchangeteam.com/archive/2010/07/09/455445.aspx
     
## Disclaimer

The sample scripts are not supported under any Microsoft standard support program or service. The sample scripts are provided AS IS without warranty of any kind. Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.