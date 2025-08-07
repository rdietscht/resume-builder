import os
from docx import Document
from docx.shared import RGBColor
from docx.shared import Pt

# Section headers to look for when scanning the document.
HEADER_OPTS = [
    "OBJECTIVE",
    "EDUCATION",
    "TECHNICAL SKILLS",
    "RELEVANT EXPERIENCE",
    "PROJECTS/RESEARCH",
    "ADDITIONAL HIGHLIGHTS",
    "AWARDS/HONORS"
]

# Currently supported resume content types.
CONTENT_TYPES = [
    "BULLETED",
    "DESCRIPTION"
]

class RRR_Parameters:

    def __init__(self):

        # Initialize included headers.
        self.included_headers = []
        for header in HEADER_OPTS:
            self.included_headers.append(header)

        # Initialize parse location.
        self.parse_path = "\\Users\\rdiet\\Documents\\Old_Resume.docx"

        # Initialize save location.
        self.save_path = "\\Users\\rdiet\\Documents\\New_Resume.docx"

    def set_parse_path(self, new_path):

        # Replace linux path style if needed.
        if ("/" in new_path):
            new_path = new_path.replace("/", "\\")

        self.parse_path = new_path
        print("Saved new parse path successfully")

    def set_save_path(self, new_path):

        # Replace linux path style if needed.
        if ("/" in new_path):
            new_path = new_path.replace("/", "\\")

        self.save_path = new_path
        print("Saved new save path successfully")

    def remove_header(self, header):

        # Ensure the user's header is a valid option and is not already removed from self.included_headers.
        if (not header in HEADER_OPTS):
            print(f"Sorry, {header} is not a valid option.")
        elif (not header in self.included_headers):
            print(f"Header {header} already removed from the header options.")
        else:
            self.included_headers.remove(header)
            print(f"Successfully removed {header} from the included header list.")

    def add_header(self, header):

        if (not header in HEADER_OPTS):
            print(f"Sorry, {header} is not a valid option.")
        elif (header in self.included_headers):
            print(f"Header {header} is already included within the header options list.")
        else:
            self.included_headers.append(header)
            print(f"Successfully added {header} to included headers list.")

    def show_included_headers(self):

        # Format included header options.
        formatted_str = "  Included Headers: (+/included, -/removed)\n"
        for header in HEADER_OPTS:
            if (header in HEADER_OPTS):
                formatted_str += '    + ' if header in self.included_headers else '    - '
                formatted_str += f"{header.capitalize()}\n"
        formatted_str += "\n"
        return formatted_str

    def show_parse_path(self):
        return f"  Resume Content Path: {self.parse_path:>47.47}\n"

    def show_save_path(self):
        return f"  Document Save Location: {self.save_path:>44.44}\n"

    def __str__(self):
        formatted_str = f"{'CURRENT PARAMETERS':<25}\n"

        formatted_str += self.show_included_headers()
        formatted_str += self.show_parse_path() + "\n"
        formatted_str += self.show_save_path() + "\n"

        return formatted_str

# Stores all the content and formatting info for each section of a single resume document.
class File_Handle:

    def __init__(self):
        self.sections = [] # sections stores Section instances


# Stores any subsections and their content.
class File_Section:

    def __init__(self, title):
        self.title = title
        self.subs = [] # for sections with no content (i.e., a single paragraph) a special "__do_not_use" section will be appended


class File_Sub:

    def __init__(self, title):
        self.title = title
        self.content_list = [] # stores array of File_Content instances


# Stores a single unit of file content, either a bulleted list or a paragraph description.
class File_Content:

    def __init__(self, content_type, content):
        self.type = content_type
        self.content = content # of type: string[] (BULLETED) | string (DESCRIPTION)


# UTILITY FUNCTIONS
def define_parameters():

    # Greet the user.
    wipe_terminal()
    print(f"{'Welcome':^25}")
    print(f"{'RAD-Resume-Retrofitter':^25}")
    print()

    # Ask to configure settings for an RRR_Parameters instance.
    params = RRR_Parameters()
    user_input = ""
    while (user_input != "N" and user_input != "NEXT"):

        # EARLY EXIT - The user has quit the program.
        if (user_input == "Q" or user_input == "QUIT"):
            print("Goodbye!")
            quit()

        # Handle modification of included headers list.
        if (user_input == "H" or user_input == "HEADER"):
            wipe_terminal()
            while(user_input != 'Q' and user_input != 'QUIT'):

                # ADDING A HEADER.
                if (user_input == "A" or user_input == "ADD"):

                    wipe_terminal()
                    print(params.show_included_headers())
                    print("Which header?")
                    print()
                    print("Type name of header: ", end="")
                    user_input = input().upper()
                    wipe_terminal()
                    params.add_header(user_input)

                    # Once a header is successfully added, return to the main menu.
                    break

                # REMOVING A HEADER.
                if (user_input == "R" or user_input == "REMOVE"):

                    wipe_terminal()
                    print(params.show_included_headers())
                    print("Which header?")
                    print()
                    print("Type name of header: ", end="")
                    user_input = input().upper()
                    wipe_terminal()
                    params.remove_header(user_input)

                    # Once a header is successfully added, return to the main menu.
                    break

                print("Would you like to add/remove a header?")
                print("(Q/Quit to go back to previous menu, A/Add to add, R/Remove to remove..)")
                print()
                print("Enter: ", end="")

                user_input = input().upper()

            print()

        # Handle changing of parse/save location.
        if (user_input == "P" or user_input == "PARSE" or user_input == "S" or user_input == "SAVE"):

            wipe_terminal()

            parse_flag = False # indicates the user is saving a new parse path

            # Save to appropriate place.
            if (user_input == "P" or user_input == "PARSE"):
                parse_flag = True

            print("Please type in the full path for the file.")
            print("(Q/Quit to cancel and go back to previous menu..)")
            print()
            print("New Path: ", end="")
            user_input = input()

            while (user_input.upper() != 'Q' and user_input.upper() != 'QUIT'):

                # Only accept non-empty path settings.
                if (user_input != ""):
                    if (parse_flag):
                        wipe_terminal()
                        params.set_parse_path(user_input)
                        break
                    else:
                        wipe_terminal()
                        params.set_save_path(user_input)
                        break

                print("Please type in the full path for the file.")
                print("(Q/Quit to cancel and go back to previous menu..)")
                print()
                print("New Path: ", end="")
                user_input = input()

            print()

        # Show current status of the parameters.
        print(params)

        print("Please Select an option:")
        print("(H/Header to modify headers, P/Parse to modify parse location, S/Save to modify save location..)")
        print("(Q/Quit to cancel, N/Next to continue with the saved parameters..)")
        print()
        print("Enter: ", end="")
        user_input = input().upper()

    return params

 
def scan_formatted_document(parse_path):

    print("Scanning document content...")
    print()

    # TODO - Performance can be improved by interpreting each line/character as it is read.
    r_lines = [] # stores a list of strings representing each line in the file.
    with open(parse_path) as fh:

        # Gather each line of the file.
        r_lines = fh.readlines()

        # Show the user content of the file scanned.
        print(f"{'RESUME CONTENT SCANNED':^100}")
        print('=' * 100)
        print()
        for r_line in r_lines:
            if (len(r_line) > 100):
                print(f"{r_line:<97.97}...")
            else:
                print(f"{r_line:<100.100}")
        print()
        print('=' * 100)
        print()

    # Go through each line, creating File_Section instances as they are encountered.
    # TODO


"""Create a .docx file using the file content and parameters.

@params
    content - A File_Content instance with the contents to be included in the formatted docx.
    params - The custom user parameters to include when formatting the document.
"""
def create_formatted_document(content, params):

    print("Creating + Formatting document...")

    # Create the Word document.
    doc = Document()

    # Construct included headers w/ section content.
    for header in params.included_headers:
        run = doc.add_heading().add_run()
        font = run.font
        font.color.rgb = RGBColor(0, 0, 0)
        font.name = 'Calibri'
        font.size = Pt(12)
        run.text = header.capitalize()

        # TODO - Write content parameter values here and transform into a bulleted list in the Word doc.

    # Print a confirmation to show the operation was successful.
    doc.save(params.save_path)
    print(f"Done - Word document saved in: {params.save_path}")
    print()


def wipe_terminal():
    # wipe terminal (either Windows/Linux)
    os.system('cls' if os.name == 'nt' else 'clear')

# Main script
if __name__ == '__main__':

    # Ask the user to define parameters (i.e., path to
    # formatted document, path to created document,
    # headers to include w/ options, etc...).
    parameters = define_parameters()

    # Obtain file information and content details. Log any errors encountered.
    sections = scan_formatted_document(parameters.parse_path)

    # Create a new word file using configured settings.
    create_formatted_document(sections, parameters)
