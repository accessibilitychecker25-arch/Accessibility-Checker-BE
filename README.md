#this gets the repo
git clone repo 

#this gets up to date code
git pull

#this creates a branch which you can work on
git checkout -b "djo/your-branch-description"

#this installs everything you need
npm i

#this gives you secrets
get .env file from DJ or put secrets in manually into .env file which you create

##VERY IMPORTANT
make sure you create a git ignore file (ask chatgpt if you have never done this before) which ignores your .env file

#this runs the program
node autotag-pdf.js

#creates a branch with your changes
git push 

#we can review pull requests as a team to identify if things are  good for merge.
