import os
import settings.conf as c


class FilePath():
    def __init__(self):
        self.cwd_path = os.getcwd()

    def get_files(self):
        try:
            if c.file_name:
                if os.path.exists(c.file_name):
                    return c.file_name
        except Exception as e:
            pass
        filelist = []
        for file in os.listdir(self.cwd_path):
            if '.xmind' in file:
                file_path = self.cwd_path + r'\%s' % file
                filelist.append(file_path)
        return filelist

filelist = FilePath()

if __name__ == '__main__':
    f = FilePath()
    f.get_files()
