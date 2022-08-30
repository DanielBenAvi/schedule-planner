class Schedual():

    def __init__(self) -> None:
        self.courses = []
        self.schedualScore = 0
        self.valid = True

    def addCourse(self, course) -> None:
        self.courses.append(course)
        self.courses.sort(key=lambda x: x.start)  # order after adding

    def calcScore(self) -> int:
        for course in self.courses:
            self.schedualScore += course.score
        return self.schedualScore

    def isValid(self) -> bool:
        maxBreak = 1 # max hours of breaks
        for index in range(0, len(self.courses)-1):
            # remove not valid
            if(self.courses[index].end > self.courses[index + 1].start):
                return False
            # remove if have breaks longer than 2 hours
            if(int(self.courses[index].start/100) == int(self.courses[index + 1].start/100)):
                if((int(self.courses[index+1].start%100) - int(self.courses[index].end%100)) > maxBreak):
                    return False

        for index in range(0, len(self.courses)):
            # remove end after 20:00
            if(self.courses[index].end%100>20):
                return False
            # remove fridays
            if(int(self.courses[index].start/100)>5):
                return False
        
        # remove days with only 2 hours
        for day in range(1,7):
            sum = 0
            for index in range(0, len(self.courses)):
                if(int(self.courses[index].start/100)==day):
                    sum = sum + int(int(self.courses[index].end%100)-int(self.courses[index].start%100))
            if(sum <= 3 and sum > 0):
                return False

        return True

    def __str__(self) -> str:
        str = '------\n'
        str += f'score = {self.schedualScore}\n'
        for course in self.courses:
            str += course.__str__() + '\n'
        return str