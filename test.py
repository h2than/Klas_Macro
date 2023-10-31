import re

class Thread():
    def __init__(self):
        self.opt = 7
        self.title = "[231321아아야야야] 테스트 성공"
        self.pattern = self.get_pattern()
        self.work_function = self.get_work_function()

    def remove(self):
        matches = self.pattern.findall(self.title)
        if matches:
            self.title = re.sub(self.pattern, '', self.title)
            self.title = self.title.strip()

    def get_pattern(self):
        re_mapping = {
            5: r'\[\d+\]',
            6: r'\[^\d]+\]',
            7: r'\[[^\]]*\]',
        }
        return re.compile(re_mapping.get(self.opt, r'\[.*\]'))

    def get_work_function(self):
        work_mapping = {
            5: self.remove,
            6: self.remove,
            7: self.remove
        }
        return work_mapping.get(self.opt)

    def display_title(self):
        print(self.title)

t = Thread()
t.work_function()
t.display_title()  # Display the modified title
