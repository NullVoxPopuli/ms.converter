# ms.converter
Ms. Converter is a REST API built on .NET for converting various Microsoft Office formats to and from HTML

# Additional Setup
Install Sass and Grunt for compiling to CSS   
http://jimfrenette.com/2015/01/visual-studio-mvc-web-app-bourbon/

# The Converter
http://powertools.codeplex.com/ installed as a NuGet

# Running
 - run `grunt` in a git power shell (optional every run after, unless scss files change)
 - click the play button in Visual Studio 2015

# Authentication

## For normal registration / non-OAuth2 accounts
 - Post to /api/Account/Register as shown here (with the Postman extension): ![Registration](http://i.imgur.com/zY5sKu0.png)

## For OAuth2 registration (Google login)


# Adding S3 credentials for doc storage, image hosting, and storing batch requests
 - coming soon

# ToDo
 - Allow users to set up a connection to Amazon S3 for image hosting for the converted HTML
 - Send docx files in batch via a zip, and return a zip of the results
 