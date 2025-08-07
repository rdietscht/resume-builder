# RAD-Resume-Retrofitter
## A simple resume-building utility

This utility is empowered with `python-docx` and is used primarily to construct custom word documents off formatted document information.

## Usage

To use this utility, simply run the main script as a python3 script:

```
python3 rrr-converter.py
```

Interacting with the script will allow you to customize and set options for your resume, as needed.

## Resume Scanner

This utility works by scanning resume content from a very particular style of text document. In order for your resume to be successfully scanned, the document used to gather the information written to the resume must match this format as closely as possible. Below are the guidelines for following the format:

* A resume is split into "sections" with each section containing a section header and section content.
* A section's content may be divided into "sub-sections" that denote a grouping of items related to the overhead section topic. A sub-section may also be used more generally to render a larger bulleted point.
* Section content may vary between either a text description (i.e., a paragraph) or a bulleted list.
* Newlines are treated as they are. If a newline is encountered, the word document will introduce a line-break. **The only exception to this rule are bullet points (*) from a bulleted list, which will _automatically_ add a line-break to render the new bullet point.**
* Each section header begins with a pound (#) and should be isolated on its own line.
* Each sub-section header within a section is surrounded by dollar sign ($).
* Bulleted list items are surrounded by a backticks (`) where each bullet point is delimited by an asterisk (*).

> Any personal info to be listed on the resume does NOT need to be described within the scanned resume document. The `rrr-converter.py` utility will automatically add this near the top of the resume after you have specified any personal contact details during settings configuration.

An example of a legitimate scanned resume template is shown below:

```plaintext
#Objective
"""Computer Science student passionate about human-centered technology and cutting-edge automation, with real-world experience building accessible and scalable web and mobile applications using React, Next.js, and TypeScript. Proven ability to take ownership of user-facing systems, collaborate in cross-functional teams, and rapidly build and iterate in fast-paced startup environments. Eager to contribute to narb’s mission by developing seamless AI-powered voice-call systems while leveraging skills in full-stack engineering, real-time APIs, and voice-integrated UX."""

#Education
$James Madison University (JMU), Harrisonburg, VA$
$Bachelor of Science in Computer Science$
$Minor in Robotics and Honors$
$May 2026$, $GPA: 3.63$
$Dean’s List Spring 2023 & Spring 2024$
$President’s List Fall 2022$
$UPE Honor Society Member$
$JMU Honors College Member$

#Technical Skills
$Languages:$ Python, JavaScript, TypeScript, Java, PHP, C, Rust, Haskell
$Frontend Technologies:$ React, React Native, Next.js, Svelte, Tailwind CSS, HTML5, CSS3
$Backend & APIs:$ Node.js, Express.js, Django, RESTful APIs, PHP
$Databases:$ PostgreSQL, MySQL, MongoDB
$DevOps & Tooling:$ Git (CLI and GitHub), Docker, CI/CD Workflows, Trello, Bash, Agile (Scrum), WebSockets
$Testing & Debugging:$ Unit Testing (Jest, JUnit, Pytest), Integration Testing, Debugging Tools, Logging, Profiling
$Design & UX:$ Accessibility (WCAG), Responsive Design, Figma
$Other:$ ROS2, Raspberry Pi, Arduino, Visual Studio Code, Linux

#Relevant Experience
$Software Consultant / Sr. Programmer, SerialByte, Short Pump, VA, Dec 2023 – Present$
`*Built full-stack web tools using React, Next.js, Tailwind, and Svelte*Collaborated directly with clients to gather requirements and iterate quickly*Designed scalable and accessible frontend systems with reusable components*Integrated RESTful APIs and implemented Git-based CI/CD workflows*Managed sprint tracking using Trello and team GitHub repos`
$Associate Software Engineer, GenXC Group, Remote, Apr 2021 – Aug 2024$`*Developed UI components and integrated backend services using Next.js, Tailwind, and JavaScript*Built responsive, WCAG-compliant UIs for internal productivity platforms*Collaborated with designers and backend engineers to optimize workflows and resolve integration bugs*Wrote unit and integration tests to support scalable deployment`
$Lead Teaching Assistant, CS Department, James Madison University, Aug 2023 – Present$`*Instructed and mentored students in foundational CS topics*Held weekly office hours supporting debugging and problem-solving in Python, Java, and C*Collaborated with faculty to identify course pain points and improve materials*Trained new TAs and contributed to inclusive classroom culture`
```