# dependancies
import docx2txt
import xml.etree.ElementTree as etree


# import document
fulltext = docx2txt.process("test.docx")

# split document into questions list
assessment = fulltext.split("ASSESSMENT QUESTIONS")
category = assessment[0].split("REVIEWERS:")
category_name = category[0].rstrip()
questions = assessment[1].split("#")

# def xml indent formating
def indent(elem, level=0):
    i = "\n" + level*"  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            indent(elem, level+1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i

# def get_correct
def get_correct(answers):
    corr = []
    for a in answers:
        if a[0] == "*":
            corr.append(a)
    return corr

# def get_answers
def get_answers(answers):
    ans = []
    for a in answers:
        if len(a) > 1:
            ans.append(a)
    return ans

# zero out xml and quiz objects, define quiz as element
xml = 0
quiz = 0
quiz = etree.Element("quiz")

# add lms category
question = etree.SubElement(quiz, "question", type="category")
category = etree.SubElement(question, "category")
etree.SubElement(category, "text").text = "$system$/"+category_name

# iterate through questions
for i in questions:
    if  len(i) > 10:
        
        # separate question string
        q_string = i.split("QUESTION:")
        q_text = q_string[1].split("ANSWERS:")
        q_answers = q_text[1].split("CORRECT FEEDBACK:")
        answers_list = q_answers[0].split("\n\n")
        correct = q_answers[1].split("IN")
        correct = correct[0].rstrip()
        incorrect = q_answers[2]
        q = q_text[0].rstrip()
        
        # call answer logic functions
        answers = get_answers(answers_list)
        corr_ans = get_correct(answers)

        # build question
        question = etree.SubElement(quiz, "question", type="multichoice")
        
        # name question
        name = etree.SubElement(question, "name")
        etree.SubElement(name, "text").text = q + "(" + str(questions.index(i)) + ")"
        
        # actual question text
        questiontext = etree.SubElement(question, "questiontext")
        etree.SubElement(questiontext, "text").text = q
        
        # feedback
        generalfeedback = etree.SubElement(question, "generalfeedback")
        etree.SubElement(generalfeedback, "text").text = ""
        
        # meta
        etree.SubElement(question, "defaultgrade").text = "1.0000000"
        etree.SubElement(question, "penalty").text = "0.0000000"
        etree.SubElement(question, "hidden").text = "0"
        if len(corr_ans) > 1:
            etree.SubElement(question, "single").text = "false"
        else:
            etree.SubElement(question, "single").text = "true"
        etree.SubElement(question, "shuffleanswers").text = "true"
        etree.SubElement(question, "answernmbering").text = "abc"
        
        # conditional feedback
        correctfeedback = etree.SubElement(question, "correctfeedback")
        etree.SubElement(correctfeedback, "text").text = correct

        partiallycorrectfeedback = etree.SubElement(question, "partiallycorrectfeedback")
        etree.SubElement(partiallycorrectfeedback, "text").text = "Your answer is partially correct."

        incorrectfeedback = etree.SubElement(question, "incorrectfeedback")
        etree.SubElement(incorrectfeedback, "text").text = incorrect
        
        #answers
        for a in answers:
            frac = str((len(corr_ans)/len(answers))*100)
            if len(corr_ans) == 1:
                if a in corr_ans:
                    answer = etree.SubElement(question, "answer", fraction="100")
                    etree.SubElement(answer, "text").text = a
                else:
                    answer = etree.SubElement(question, "answer", fraction="0")
                    etree.SubElement(answer, "text").text = a
            else:
                if a in corr_ans:
                    answer = etree.SubElement(question, "answer", fraction=frac)
                    etree.SubElement(answer, "text").text = a
                else:
                    answer = etree.SubElement(question, "answer", fraction="-"+frac)
                    etree.SubElement(answer, "text").text = a
        
# call formatter, export to file
xml = etree.ElementTree(quiz)
indent(quiz)
xml.write("test.xml", encoding='utf-8', xml_declaration=True) # ENCODING MUST BE UTF-8, DECLARATION MUST BE TRUE