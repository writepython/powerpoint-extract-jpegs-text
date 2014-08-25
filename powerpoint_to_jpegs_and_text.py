import sys
import os
import string

def printable_escape(s):
    acceptable_chars = string.printable[:-3]
    return "".join(c for c in s if c in acceptable_chars)

def powerpoint_to_text(document, txt_filepath):
    f = open(txt_filepath, 'w')
    titles_and_text = []
    title_format_string = '"%s"'
    pages = document.getDrawPages()
    num_of_pages = pages.getCount()

    for page_num in range(num_of_pages):
        current_page = pages.getByIndex(page_num)
        page_name_unescaped = current_page.Name
        page_name = printable_escape(page_name_unescaped)
        titles_and_text.append(title_format_string % page_name)
        num_of_shapes = current_page.getCount()
        for shape_num in range(num_of_shapes):
            current_shape = current_page.getByIndex(shape_num)
            try:
                shape_text_unescaped = current_shape.getText().getString()
                shape_text = printable_escape(shape_text_unescaped)
                titles_and_text.append(shape_text)
            except:
                pass
    file_text = '\n<br>\n'.join(titles_and_text)
    f.write(file_text)
    f.close()

def powerpoint_to_jpegs(powerpoint_filepath, output_dir, extract_text=False):
    import ooutils
    from com.sun.star.beans import PropertyValue
    import uno
            
    oor = ooutils.OORunner()
    desktop = oor.connect()

    html_property = (PropertyValue("FilterName" , 0, "impress_html_Export", 0),)
    input_url = uno.systemPathToFileUrl(powerpoint_filepath) 
    document = desktop.loadComponentFromURL(input_url, "_blank", 0, ())
    document.storeToURL("file://"+output_dir, html_property)

    if extract_text:
        new_txt_filepath = powerpoint_filepath.rsplit('.')[0] + '.txt' 
        powerpoint_to_text(document, new_txt_filepath)
    document.dispose()

    output_dir_files = os.listdir(output_dir)
    for filename in output_dir_files:
        if filename.endswith('.html'):
            os.remove(os.path.join(output_dir, filename))
    oor.shutdown()

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print "Usage: python powerpoint_to_jpegs.py PATH_TO_POWERPOINT_FILE OUTPUT_DIRECTORY extract_text"
    else:
        powerpoint_filepath = sys.argv[1]
        output_dir = sys.argv[2]
        try:
            extract_text = sys.argv[3]
        except:
            extract_text = False
        powerpoint_to_jpegs(powerpoint_filepath, output_dir, extract_text)
