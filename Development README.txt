For at kunne bygge en MSI pakke fra VS skal alle disse komponenter lægges ind i DIN visual studio.
Filerne skal ligge her: C:\Program Files\Microsoft Visual Studio 8\SDK\v2.0\BootStrapper\Packages
under de respektive mapper.


- Der skal installeres visual studio 2005
- Der skal installeres outlook 2003
- Der skal installeres visual studio tools 2005 se for office
- Der skal installeres/kopieres office PIA redist. versioner ind
- Der skal installeres/kopieres VSTO 2005 SE redist. Ind

Office 2003 PIA redist findes her... http://www.microsoft.com/downloads/details.aspx?familyid=3c9a983a-ac14-4125-8ba0-d36d67e0f4ad&displaylang=en

Office 2007 PIA redist findes her...
http://www.microsoft.com/downloads/details.aspx?familyid=59daebaa-bed4-4282-a28c-b864d8bfa513&displaylang=en

Begge filer skal pakkes ud og lægges i mappen

VSTO 2005 SE x86 er her
http://www.microsoft.com/downloads/details.aspx?familyid=f5539a90-dc41-4792-8ef8-f4de62ff1e81&displaylang=en

VSTO2005se installer samples
http://www.microsoft.com/downloads/details.aspx?familyid=6991E869-8D5B-45F4-91E7-B527BD236F4C&displaylang=en

Denne installation laver et dir der hedder:
C:\Program Files\Microsoft Visual Studio 2005 Tools for Office SE Resources\VSTO2005SE Windows Installer Sample

Under dette dir skal de andre downloadede ting lægges ind.
Når de kickstartpakker der skal bruges er komplette (filerne kopieret ind) skal de så kopieres til visual studio. Dette skal kun gøres én gang.
Pakkerne skal ligge her i visual studio.
C:\Program Files\Microsoft Visual Studio 8\SDK\v2.0\BootStrapper\Packages Det er her alle prereqs ligger.


---------------------------------------------------------------------
Vedr. Inno Setup delen

I deployment mappen ligger de inno setup scrips der skal anvendes. 
Alle de filer der skal med i installeren skal ligge i deployment mappen. Compile derefter inno scripts og din installer ligger i output mappen i deployment.

Der skal være lavet MSI pakker fra visual studio før inno installeren laves. Når MSI pakkerne er lavet skal de kopieres til deployment mappen.

----------------------------------------------------------------------
----------------------------------------------------------------------
SUPPORT INFORMATION

PÅ VISTA KAN MAN INSTALLERER FØLGENDE VSTO SE OPDATERING VED PROBLEMER MED AT SPAMGRABBER FORSVINDER.

http://www.microsoft.com/downloads/details.aspx?FamilyId=607D2E96-31F9-4FD5-A888-DEC4FC2D67AB&displaylang=en
