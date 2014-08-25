def powerpoint_file_to_jpegs_and_text(self, powerpoint_filepath, output_dir):
    import os
    import subprocess

    oo_program_path = '/opt/openoffice.org3/program/'
    oo_python = oo_program_path + 'python '
    ppt_to_jpeg = oo_program_path + 'powerpoint_to_jpegs_and_text.py '
    
    cmd = oo_python + ppt_to_jpeg + output_dir+powerpoint_filepath + ' ' + output_dir + ' True'
    retcode = subprocess.call(cmd, shell=True)

    image_num_names_list = [(int(f.lstrip('img').rstrip('.jpg')), f) for f in os.listdir(output_dir) if f.startswith('img') and f.endswith('.jpg')]
    image_num_names_list.sort()
    image_names_list = [tup[1] for tup in image_num_names_list]

    powerpoint_text = open(output_dir + powerpoint_filepath.rsplit('.')[0] + '.txt').read()

    return image_names_list, powerpoint_text, str(retcode)
