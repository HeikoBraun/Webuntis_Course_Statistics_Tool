#!/usr/bin/env python3
import datetime
import re
import tomllib

import reportlab
import reportlab.lib
import reportlab.lib.colors
import reportlab.lib.pagesizes
import reportlab.lib.units
import reportlab.platypus
import reportlab.rl_config

import webuntis

"""https://python-webuntis.readthedocs.io/en/latest/"""


class Course:
    def __init__(self, name):
        self.name = name
        self.regular = 0
        self.irregular = 0
        self.cancelled = 0
        self.alternative = 0

    def __str__(self):
        self.calc()
        return f"{self.name}: {self.num}"

    def __add__(self, other):
        self.regular += other.regular
        self.irregular += other.irregular
        self.cancelled += other.cancelled
        self.alternative += other.alternative
        return self

    @property
    def target(self):
        return self.regular + self.cancelled

    @property
    def percentU(self):
        target = self.target
        if target == 0:
            return 0
        return round(100 * (self.regular + self.irregular) / target)

    @property
    def percent(self):
        target = self.target
        if target == 0:
            return 0
        return round(100 * (self.regular + self.irregular + self.alternative) / target)

    def incr(self, name=None):
        if name is None or name == "None":
            self.regular += 1
        elif name == "irregular":
            self.irregular += 1
        elif name == "cancelled":
            self.cancelled += 1
        elif name == "alternative":
            self.alternative += 1
        else:
            print(f"Error: '{name}' as code is unknown to me.")
            exit(1)

    def get_table_entry(self):
        return [
            self.name,
            self.regular,
            self.irregular,
            self.alternative,
            self.cancelled,
            self.target,
            self.percentU,
            self.percent,
        ]


def has_same_timeslot(lesson_1, lesson_2):
    return lesson_1.start == lesson_2.start and lesson_1.end == lesson_2.end


def write_initial_toml():
    with open("config.toml", "w") as fd:
        fd.write(
            r"""server="webuntis-server"      # e.g. "herakles.webuntis.com"
school="webuntis-school-name" # e.g. "FannyLGym"
username="username"
password="password"
# useragent is optional, but should be given with mail for fairness
# so that an administrator knows whom to contact!
useragent="WebUntis Test (yourmail@example.com)"
# classes is optional
# 1. String => only this class will be used. e.g. "9b"
# 2. List of strings => these classes will be used. e.g. ["9b", "9d"]
# 3. not given => all classes matching \d+[a-z] will be used
classes="9b"
"""
        )
    print(
        "Wrote out an example configuration file 'config.toml'.\n"
        "Please fill out the relevant values."
    )


def get_lesson_name(lesson):
    if lesson.activityType == "Unterricht":
        if lesson.subjects:
            ret = str(lesson.subjects[0])
        else:
            print(f"Warning! {lesson}")
            ret = ""
        if lesson.studentGroup:
            ret = ret + "/" + str(lesson.studentGroup)
    else:
        ret = lesson.lstext
    if not ret:
        ret = "unbekannt"
    return ret


def myFirstPage(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica-Bold", 16)
    canvas.drawCentredString(
        reportlab.rl_config.defaultPageSize[0] / 2.0,
        reportlab.rl_config.defaultPageSize[1] - 60,
        doc.title,
    )
    canvas.setFont("Helvetica-Bold", 14)
    canvas.drawCentredString(
        reportlab.rl_config.defaultPageSize[0] / 2.0,
        reportlab.rl_config.defaultPageSize[1] - 90,
        doc.date_period,
    )
    canvas.restoreState()


def gen_pdf(courses, class_name, school_year, alternatives_used):
    filename = f"Unterrichtsstatistik_für_Klasse_{class_name}.pdf"
    pdf = reportlab.platypus.SimpleDocTemplate(
        filename=filename,
        pagesize=reportlab.lib.pagesizes.A4,
        title=f"Unterrichtsstatistik für Klasse {class_name}",
        topMargin=4 * reportlab.lib.units.cm,
    )
    pdf.date_period = f"Anfang Schuljahr {school_year} bis {datetime.date.today()}"

    headings = [
        "Fach",
        "regulär",
        "zusätzlich",
        "alternativ",
        "ausgefallen",
        "soll",
        "ProzentU",
        "ProzentA",
    ]

    data = [headings]
    all_courses = Course("")
    for course in sorted(courses, key=lambda x: x.name):
        data.append(course.get_table_entry())
        all_courses += course
    data.append(all_courses.get_table_entry())

    table = reportlab.platypus.Table(data=data)
    style = [
        # header
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        # data
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), 10),
        ("INNERGRID", (0, 0), (-1, -2), 0.25, reportlab.lib.colors.black),
        ("BOX", (0, 0), (-1, -2), 0.5, reportlab.lib.colors.black),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        # total (last line)
        ("INNERGRID", (1, -2), (-1, -1), 0.25, reportlab.lib.colors.black),
        ("BOX", (1, -1), (-1, -1), 0.5, reportlab.lib.colors.black),
    ]

    # loop is just to make a grey background on each second line for better readability
    grey = reportlab.lib.colors.Color(0.8, 0.8, 0.8)
    for i in range(len(data)):
        if i % 2 == 1:
            style.append(("BACKGROUND", (0, i), (-1, i), grey))
    table.setStyle(reportlab.platypus.TableStyle(style))

    # alternatives table
    data = [["Alternativ", "Stunden"]]
    data.extend(
        [k, v]
        for k, v in sorted(
            alternatives_used.items(), key=lambda item: item[1], reverse=True
        )
    )
    table_a = reportlab.platypus.Table(data=data)
    style_a = [
        # header
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        # data
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), 10),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, reportlab.lib.colors.black),
        ("BOX", (0, 0), (-1, -1), 0.5, reportlab.lib.colors.black),
    ]
    table_a.setStyle(reportlab.platypus.TableStyle(style_a))

    pdf.build([table, reportlab.platypus.PageBreak(), table_a], onFirstPage=myFirstPage)
    print(f"Written {filename}")


def work_on(config):
    with webuntis.Session(
        username=config["username"],
        password=config["password"],
        server=config["server"],
        school=config["school"],
        useragent=config.get("useragent", "WebUntis test"),
    ).login() as session:
        actual_school_year = session.schoolyears().filter(is_current=True)[0]
        end_date = actual_school_year.end
        if end_date > datetime.datetime.now():
            end_date = datetime.datetime.now()
        classes = config.get("classes", [])
        if classes:
            if isinstance(classes, str):
                classes = [classes]
            elif isinstance(classes, list):
                pass
            else:
                print(f"Error: {classes}")
                exit(1)
        else:
            classes = [
                c.name for c in session.klassen() if re.match(r"\d+[a-z]", c.name)
            ]
        for class_name in classes:
            klasse = session.klassen().filter(name=class_name)[0]
            lessons = session.timetable_extended(
                klasse=klasse, start=actual_school_year.start, end=end_date
            )
            courses = {}
            alternatives = [
                lesson for lesson in lessons if lesson.activityType != "Unterricht"
            ]
            alternatives_used = {}
            for lesson in lessons:
                course_name = get_lesson_name(lesson)
                if lesson.activityType != "Unterricht":
                    continue
                if course_name not in courses:
                    courses[course_name] = Course(course_name)
                courses[course_name].incr(str(lesson.code))
                if lesson.code == "cancelled":
                    alternatives_found = [
                        l for l in alternatives if has_same_timeslot(lesson, l)
                    ]
                    if alternatives_found:
                        courses[course_name].incr("alternative")
                        alternative_name = get_lesson_name(alternatives_found[0])
                        alternatives_used[alternative_name] = (
                            alternatives_used.get(alternative_name, 0) + 1
                        )
            # print(alternatives_used)
            gen_pdf(
                courses.values(), class_name, actual_school_year.name, alternatives_used
            )


if __name__ == "__main__":
    try:
        with open("config.toml", "rb") as f:
            toml_config = tomllib.load(f)
    except FileNotFoundError:
        print("Error: configuration file was not found.")
        write_initial_toml()
        exit(1)
    work_on(toml_config)
