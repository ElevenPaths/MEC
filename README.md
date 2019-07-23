# Masked Extension Control (MEC)

<img src="https://www.elevenpaths.com/wp-content/uploads/2014/03/11paths_logo.png?v=3&s=200" title="LogoApp" alt="LogoApp">

Masked Extension Control (MEC) Masked Extension Control is a program that makes your Windows rely on magic numbers, and not only on extensions, to choose the program that will be used to open the file. This is much safer for your system, since a lot of attacks begin by fooling extensions and trying to be opened or executed by a vulnerable program, instead of the one the file is really supposed to be opened with.

## Developer Mod.

Please refer to https://mec.e-paths.com for more infomation.

### Prevent attacks based on fake extensions
Attackers usually change extensions in files to make you trust the file, and this is dangerous. For example, some very popular attacks make RTF files to be opened with Word, just by modifying the RTF extension to DOC or DOCX. This way, they build exploitable RTF files that will take advantage of Word vulnerabilities or weaknesses to release their payload. However, if these RTF files are opened with WordPad, the threat will disappear.

### Easy to use
This program does not need to be resident on memory. It modifies Windows registry to open MHT, DOC, RTF and DOCX files with the program that they should be opened, so trusting in magic number instead of extensions. If you want to stop using it, you just have to uninstall it.

### Most common formats and extensions
Not only RTF and DOC files but .MHT files as well: if they are opened with Word, some vulnerabilities may be exploited, but if they are opened with a browser, is less likely that something happens. Masked Extension Control works even with malformed magic numbers in RTF (which is much more common than you may think).

### Prerequisites

Office 365 - Office 2010

Please remember to add NLog package before compiling. Add it as a NuGet package from the project itself.

## Languajes y Technologies

* C#

## Authors

 *Jose Sperk** 
 *Sergio De los Santos** 
