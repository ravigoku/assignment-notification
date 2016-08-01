
## README
Every exam, every quiz, every project, and every assignment is roughly scheduled in the beginning semester for every class. This is atleast true if you have a half decent teacher. If you are graced with having a decent teacher, he or she may even hand it out on paper during class or upload it online.

This program a tool for a 'user' who tracks all assignments due dates of multiple 'cleints', to help make sure 'clients' are organized and keeping up with their assignments. For each 'client', the program reads a spreadsheet containing due dates for exams, quizzes, projects, and weekly assignments. The program checks if there is an exam/quiz/project/assignment coming up in the next week, and uses a twilio REST API to message the client of the exam, requesting that he or she start organizing their "study approach." Every Sunday, if the program is executed, a list of all person's upcoming assignments in the next week will be messaged to the user.

