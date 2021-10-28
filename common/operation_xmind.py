import xmind


class ReadXmind():
    def __init__(self, index=None, file_path=None):
        self.workbook = xmind.load(file_path)
        self.sheets = self.workbook.getSheets()
        self.all_sheet, self.first_title = self.get_sheet_element_list(index)

    def get_sheet_element_list(self, index):
        all_sheet = []
        first_title = ''
        for sheet in self.sheets:
            root_topic = sheet.getRootTopic()
            first_title = root_topic.getTitle()
            test_list = []
            self.node_recursion(root_topic, test_list, '')
            all_sheet.append(test_list)
        if index == None:
            return all_sheet[0], first_title
        else:
            if isinstance(index, int):
                return all_sheet[index], first_title
            elif isinstance(index, list):
                sheet_list = []
                for i in index:
                    sheet_list.append(all_sheet[i])
                return sheet_list, first_title
            else:
                raise TypeError("The object ReadXmind() 1 takes is not vaild\tcode : ParamsError")

    def node_recursion(self, root, test_list, topic_str):
        topics = root.getSubTopics()
        sub_topic_count = sum(1 for _ in topics)
        if root.getTitle():
            topic_str = topic_str + '\n' + root.getTitle()
        else:
            sub_topic_count = -1
        if sub_topic_count > 0:
            for elem in topics:
                self.node_recursion(elem, test_list, topic_str)
        else:
            test_list.append(topic_str.strip().split('\n'))


if __name__ == '__main__':
    r = ReadXmind()
    all_sheet = r.all_sheet
    print(all_sheet)
