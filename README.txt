
Extracts images from a PowerPoint file for display on a website or similar.  Additionally, extracts any text from the PowerPoint file to improve search results.

== Usage ==

<python> powerpoint_to_jpegs_and_text.py <path_to_powerpoint_file> <output_directory> <extract_text>

<python> is the version of Python supplied with OpenOffice. See OpenOffice Configuration below.
<path_to_powerpoint_file> is the fully qualified path to a PowerPoint file.
<output_directory> is the fully qualified path to an output directory for extracted text and images.
<extract_text> If this third argument is present, text will be extracted from the PowerPoint and written to the same output directory as the images.

To interact with powerpoint_to_jpegs_and_text.py from within a Zope process running a different verion of python, use zope_ext_method.py

== Dependencies ==

OpenOffice: http://download.openoffice.org/other.html#en-US

== OpenOffice Configuration ==

1) Edit the two OpenOffice config files changing "ooSetupInstCompleted" to true, otherwise OpenOffice will hang.

On CentOS they are located at:
/opt/openoffice.org3/basis-link/share/registry/data/org/openoffice
/opt/openoffice.org/basis3.1/share/registry/data/org/openoffice

The xml should look like this when you're done:
    <prop oor:name="ooSetupInstCompleted">
      <value>true</value>
    </prop>

2) Run OpenOffice in headless mode and listening on a port.

Locate the OpenOffice program directory.  On CentOS it is located at: /opt/openoffice.org3/program/

Execute the soffice command:
/opt/openoffice.org3/program/soffice "-accept=socket,host=localhost,port=2002;urp;" -headless -nofirststartwizard -norestore -nologo

Now OpenOffice should be running on port 2002 and waiting for requests.  You can run netstat -lnp to check this.

3) Make sure you can connect to interact with PowerPoint files in OpenOffice using the version of Python supplied with OpenOffice.

/opt/openoffice.org3/program/python
>>> powerpoint_filepath = <path_to_your_ppt_file>
>>> import uno
>>> local = uno.getComponentContext()
>>> resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
>>> context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
>>> desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
>>> document = desktop.loadComponentFromURL("file://"+powerpoint_filepath, "_blank", 0, ())

== Resolving Issues ==

If you're having issues connecting to OpenOffice or getting it to listen on a port:

1. Try adjusting BOTH assignments of localhost (soffice command and resolve method call) to 0 (zero) rather than localhost .
2. Run "pkill soffice" 

Once you've got OpenOffice up and running, and starting and stopping properly, you shouldn't have to do this anymore.
