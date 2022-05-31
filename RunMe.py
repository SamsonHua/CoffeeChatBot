#==============================================#
#                ATTA FIKA BOT                 #
#==============================================#
#       DEVELOPED BY CONRAD B & SAMSON H       #
#==============================================#
#    YOU MUST LOGIN TO FIKA@ATTABOTICS.COM     #
#               FOR THIS TO WORK               #
#==============================================# 

#Dependencies
from kivy.app import App
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.image import Image
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.core.window import Window
from kivy.config import Config
from kivy.app import App
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
import numpy
import Headers.emailFunctions as emailFunctions
import Headers.csvProcessing as csvProcessing

#============================================================#
# >>> BUILD WINDOWED APPLICATION
#============================================================#

class ATTAFika(App):
    
    def build(self):
        self.window = GridLayout()
        self.window.cols = 1
        self.window.size_hint = (0.6,0.9)
        self.window.pos_hint = {"center_x": 0.5, "center_y":0.5}
        self.icon = 'Headers/AttaLogo.ico'
        self.title = 'ATTA Fika v1.0.1'
       
        #Header Image
        self.window.add_widget(Image(source="Headers/header.png",size_hint = (1.5,1.5)))

        #Text Label
        self.welcome = Label(text="This Week's Pairings:")
        self.window.add_widget(self.welcome)

        #Pairings Preview
        layout = GridLayout(cols=1, padding=5, spacing=5,
                size_hint=(None, None), width=500)

        layout.bind(minimum_height=layout.setter('height'))

        Pairings = csvProcessing.getPairings()

        for rows in Pairings:
            if rows[0] == "NULL" and rows[1] == "NULL" and rows[2] == "NULL":
                continue
            if rows[2] == "NULL":
                name1 = emailFunctions.parseFirstName(rows[0])
                name2 = emailFunctions.parseFirstName(rows[1])
                btn = Button(text=name1 + " & " + name2, size=(580, 40),
                         size_hint=(None, None))
            else:
                name1 = emailFunctions.parseFirstName(rows[0])
                name2 = emailFunctions.parseFirstName(rows[1])
                name3 = emailFunctions.parseFirstName(rows[2])
                btn = Button(text=name1 + " & " + name2 + " & " + name3, size=(580, 40),
                         size_hint=(None, None))
            layout.add_widget(btn)
            
        # create a scroll view, with a size < size of the grid
        list = ScrollView(size_hint=(None, None), size=(600, 200),
                pos_hint={'center_x': .5, 'center_y': .5}, do_scroll_x=False)
        list.add_widget(layout)
        self.window.add_widget(list)
    
        #Generate Pairings Button
        self.button = Button(text="Send Email Invite To Pairings", size_hint = (1,0.5))
        self.button.bind(on_press=self.callback)
        self.window.add_widget(self.button)

        #Set window size
        Config.set('graphics', 'width', '1000')
        Config.set('graphics', 'height', '750')
        Config.set('graphics', 'resizable', False)
        Config.write()

        return self.window

    #============================================================#
    # >>> FUNCTION WHEN BUTTON IS PRESSED
    #============================================================#  

    def callback(self, instance):
        self.welcome.text = "Meeting Invitations have been sent! Please restart application if you need to send emails again"
        self.window.remove_widget(self.button)
        emailFunctions.sendAllEmails()

#============================================================#
# >>> RUN PYTHON APPLICATION
#============================================================#

if __name__ == "__main__":
    ATTAFika().run()