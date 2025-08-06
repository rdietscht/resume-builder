import os
from docx import Document

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

class RRR_Parameters:

    def __init__(self):
        self.included_headers = []
        for header in HEADER_OPTS:
            self.included_headers.append(header)

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

    def __str__(self):
        formatted_str = f"{'CURRENT PARAMETERS':<25}\n"

        formatted_str += self.show_included_headers()

        return formatted_str



class File_Content:
    # TODO - define this
    pass

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

        # Show current status of the parameters.
        print(params)

        print("Please Select an option:")
        print("(Q/Quit to cancel, H/Header to modify headers, N/Next to continue..)")
        print()
        print("Enter: ", end="")
        user_input = input().upper()

 

def scan_formatted_document():
    # TODO - Scan a document with a custom formatting to be able to successfully parse info
    print("formatting document...")

"""Create a .docx file using the file content and parameters.

@params
    content - A File_Content instance with the contents to be included in the formatted docx.
    params - The custom user parameters to include when formatting the document.
"""
def create_formatted_document(content, params):
    # TODO - Use the parameters and information parsed to create a word doc with python-docx package
    print("creating document...")

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
    content = scan_formatted_document()

    # Create a new word file using configured settings.
    create_formatted_document(content, parameters)
