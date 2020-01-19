import cmd
import os
class CLI(cmd.Cmd):
    def __init__(self):
        cmd.Cmd.__init__(self)
        self.prompt = "Miller2 > "  # define command prompt

    def do_dir(self, arg):
        if not arg:
            self.help_dir()
        elif os.path.exists(arg):
            print("\n".join(os.listdir(arg)))
        else:
            print( "No such pathexists.")

    def help_dir(self):
        print("syntax: dir path -- displaya list of files and directories")

    def do_quit(self, arg):
        return True

    def help_quit(self):
        print("syntax: quit -- terminatesthe application")

    # define the shortcuts
    do_q = do_quit

if __name__ == "__main__":
    cli = CLI()
    cli.cmdloop(intro="welcome to axiba")