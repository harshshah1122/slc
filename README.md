# SLC
This project is solely a desktop-based program and lacks a web-based data server, thus restricting multi user access.
The code can work best when all users operate on the same computer; Else, will result in mismatched data. 

Basics:
This project was developed in Pycharm by Harsh Shah and Dhruva Devabhaktuni.
The UI is based on the tkinter library in Pycharm.
All the data is stored in separate Excel files that are provided in the Project Files.

Process: 
To register, click on the new user option and fill out your details.
If you are returning or just have created an account, put in your details in the login fields and click enter to login.
If you are admin, use username: "admin@gmail.com" and password:"admin" to login to the admin dashboard.

User Dashboard:
The Home Page shows the personal information
The Events page displays upcoming events
The Participation page has multiple buttons. Click them will generate a code for the respective event.
The Prizes page shows the user's score, the leader's score, and the available prizes.

Admin Dashboard:
Clicking on each event will authorize the admin to enter the code stated by the user to check them into the respective event. 
The Final Winner button allows the admin to see the person at first place, second place, and third place.
All the attendance will be tracked in the EventAttendance.xlsx file and the 'Winner' column of Book1.xlsx

Control Widgets used: Buttons, Entry, Spinbox, Dropdown, Error Box

