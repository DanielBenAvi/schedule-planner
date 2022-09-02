class Course():
    def __init__(self, name, start, end, lecturer, score) -> None:
        self.name = name
        self.start = start
        self.end = end
        self.lecturer = lecturer
        self.score = score

    def __str__(self) -> str:
        weekday = {1:'Sun',2:'Mon',3:'Tue',4:'Wed',5:'Thu',6:'Fri'}
        return f'{weekday[int(self.start/100)]}: {int(self.start%100)}:00-{int(self.end%100)-1}:50 ->  {self.name} ,   {self.lecturer.__str__()},   score = {self.score}'