import os, sys, textwrap

#########################
### WINDOW DIMENSIONS ###
#########################

def window ():
    env = os.environ
    def ioctl_GWINSZ(fd):
        try:
            import fcntl, termios, struct, os
            cr = struct.unpack('hh', fcntl.ioctl(fd, termios.TIOCGWINSZ,
        '1234'))
        except:
            return
        return cr
    cr = ioctl_GWINSZ(0) or ioctl_GWINSZ(1) or ioctl_GWINSZ(2)
    if not cr:
        try:
            fd = os.open(os.ctermid(), os.O_RDONLY)
            cr = ioctl_GWINSZ(fd)
            os.close(fd)
        except:
            pass
    if not cr:
        cr = (env.get('LINES', 25), env.get('COLUMNS', 80))

    return cr[1], cr[0]

width, height = window ()

#################
### FUNCTIONS ###
#################

def printMessage (message):
    message = message.split()
    lines = []

    line = ""

    while len(message) > 0:

        if len(line) + len(message[0]) < 80:
            line = line + message[0] + " "
            message.pop(0)

            if len(message) == 0:
                lines.append(line)

        else:
            lines.append(line)
            line = ""

    for l in lines:
        print(l)

#############################
### EXTERNAL DEPENDENCIES ###
#############################

attempted = False

try:
    import docx
except ImportError:
    # on unsuccessful import...
    printMessage("You don't have the \"python-docx\" package installed, which is required for this script. You can use PIP to install it, using the command \"pip3 install python-docx\".")

    if not attempted:
        # if the script has not attempted to install the dependency already...
        response = None
        while response == None:
            print("Would you like us to try to install it?")
            print("=" * width)
            response = input(" $ ").lower()
            print("=" * width)
            if response in ["yes", "y"]:
                os.system("pip3 install python-docx")
                os.system("clear")
                attempted = True
            elif response in ["no", "n"]:
                print("Goodbye.")
                print("=" * width)
                sys.exit()
            elif response == "":
                print("Goodbye.")
                print("=" * width)
                sys.exit()
            else:
                print("Please enter yes or no.")
                response = None
    else:
        # if the script has already attempted to install the dependency...
        printMessage("We attempted to install the \"python-docx\" package but were unable to do so. Please run the command yourself.")
        sys.exit()
        print("=" * width)

attempted = False

try:
    import xlsx
except ImportError:
    # on unsuccessful import...
    printMessage("You don't have the \"py-xlsx\" package installed, which is required for this script. You can use PIP to install it, using the command \"pip3 install py-xlsx\".")

    if not attempted:
        # if the script has not attempted to install the dependency already...
        response = None
        while response == None:
            print("Would you like us to try to install it?")
            print("=" * width)
            response = input(" $ ").lower()
            print("=" * width)
            if response in ["yes", "y"]:
                os.system("pip3 install py-xlsx")
                os.system("clear")
                attempted = True
            elif response in ["no", "n"]:
                print("Goodbye.")
                print("=" * width)
                sys.exit()
            elif response == "":
                print("Goodbye.")
                print("=" * width)
                sys.exit()
            else:
                print("Please enter yes or no.")
                response = None
    else:
        # if the script has already attempted to install the dependency...
        printMessage("We attempted to install the \"py-xlsx\" package but were unable to do so. Please run the command yourself.")
        print("=" * width)
        sys.exit()

###################
### MAIN SCRIPT ###
###################

os.system("clear")

printMessage("Enter the filename and extension for your awards spreadsheet. It must be in this directory. Or, drag and drop the file from anywhere on your Mac into the Terminal.")

print("=" * width)
awardsFile = input(" $ ").strip().replace("\\", "")
print("=" * width)

if awardsFile == "":
    print("Goodbye.")
    print("=" * width)
    sys.exit()

if not os.path.isfile(awardsFile):
    printMessage("That file does not exist. Make sure you entered the name correctly and that it is in this directory.")
    print("=" * width)
    sys.exit()

awardsFileExtension = os.path.splitext(os.path.realpath(awardsFile))[1]

if awardsFileExtension not in [".csv", ".xlsx"]:
    printMessage("The awards spreadsheet must be in either CSV or XLSX format.")
    print("=" * width)
    sys.exit()

if not os.path.isfile("template.docx"):
    printMessage("We can't find the template file for your awards [\"template.docx\"].")
    print("=" * width)
    sys.exit()

if awardsFileExtension == ".csv":

    f = open(awardsFile)

    lines = f.readlines()
    lines = [l.strip().split(",") for l in lines]

    f.close()

elif awardsFileExtension == ".xlsx":

    book = xlsx.Workbook(awardsFile)

    sheetName = False
    for sheet in book:
        if not sheetName:
            sheetName = sheet.name

    sheet = book[sheetName]
    rows = sheet.rows()

    lines = []

    for row in rows:
        line = [c.value for c in rows[row]]
        lines.append(line)

l = lines[0]

if l[0].lower() in ["committee", "position", "delegate", "delegation", "award"] or l[1].lower() in ["committee", "position", "delegate", "delegation", "award"] or l[2].lower() in ["committee", "position", "delegate", "delegation", "award"]:
    response = None
    while response is None:
        printMessage("It looks like the first row of your awards spreadsheet may be a header. Should we make an award for the first row?")
        print("=" * width)
        response = input(" $ ").lower()
        print("=" * width)
        if response in ["yes", "y"]:
            pass
        elif response in ["no", "n"]:
            lines.pop(0)
        elif response == "":
            print("Goodbye.")
            print("=" * width)
            sys.exit()
        else:
            print("Please enter yes or no.")
            response = None

response = None
while response is None:
    printMessage("This script needs your awards spreadsheet to be ordered \"committee\", \"award\", \"delegation\". Is it in this order?")
    print("=" * width)
    response = input(" $ ").lower()
    print("=" * width)
    if response in ["yes", "y"]:
        pass
    elif response in ["no", "n"]:
        print("Please put it in this order.")
        print("=" * width)
        sys.exit()
    elif response == "":
        print("Goodbye.")
        print("=" * width)
        sys.exit()
    else:
        print("Please enter yes or no.")
        response = None

if not os.path.isdir("exports"):
    os.makedirs("exports")

i = 0

try:
    for l in lines:
        committee = l[0]
        award = l[1]
        delegation = l[2]

        doc = docx.Document("template.docx")

        fields = doc.tables[0]

        fields.rows[1].cells[0].paragraphs[0].text = delegation
        fields.rows[3].cells[0].paragraphs[0].text = award
        fields.rows[5].cells[0].paragraphs[0].text = committee

        filename = "__".join([committee, award, delegation])
        filename = filename + ".docx"

        doc.save("exports/" + filename)

        i = i + 1
except:
    print("The script encountered a problem.")
    print("=" * width)
    sys.exit()

print("Your awards have been successfully created.")
print(str(i) + " files created.")
print("=" * width)
