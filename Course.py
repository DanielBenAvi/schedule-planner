class Course():
    def __init__(self, name, start, end, lecturer, score) -> None:
        self.name = name
        self.start = start
        self.end = end
        self.lecturer = lecturer
        self.score = score

    def __str__(self) -> str:
        return f'{int(self.start)} - {int(self.end)}:  {self.name} ,   {self.lecturer.__str__()},   score = {self.score}'