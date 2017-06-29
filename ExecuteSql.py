import sublime
import sublime_plugin
import os
import sys
import subprocess
import errno
from timeit import default_timer as timer
import io
import codecs

# Note:
# I tried executing the query with via the Windows Shell/console and get the output via stdout,
# but I didn't figure out how to retrieve the console encoding with which I need to decode the output (and then re-encode in Unicode).
# It is much easier to execute a command with Powershell, write its output to a temp file with a given encoding
# and read it from there. Plus, you can load up the results in a tab if needed.
# The only catch is that the Powershell command "Out-File" saves a BOM with UTF8 and UTF16.
# This requires a bit of file handling to get around that.

def show_msg(msg):
    sublime.status_message("Executing SQL: {}".format(msg))

class ExecuteSqlCommand(sublime_plugin.WindowCommand):
    def run(self):
        window = self.window
        view = window.active_view()
        if view != None:
            filename = view.file_name()
            syntaxFile = view.settings().get('syntax')
            # SQL files only.
            if "SQL" in syntaxFile.upper():
                if filename == None or filename == "":
                    show_msg("current view has never been saved on disk.")
                else:
                    show_msg("executing.")
                    view.run_command("save")
                    sublime.set_timeout_async(lambda: execute_sql(window, filename), 0)
            else:
                show_msg("current view is not an SQL file.")
        else:
            show_msg("there's no view currently opened.")

def execute_sql(window, filename):
    # Settings
    settings = sublime.load_settings("ExecuteSql.sublime-settings")
    if settings is None:
        show_msg("settings aren't specified.")
        return
    
    inputPath = filename
    outputDir = get_output_dir()
    outputPath = os.path.join(outputDir, os.path.basename(filename) + ".results")

    server = settings.get("server")
    database = settings.get("database")
    showResultsInPanel = settings.get("showResultsInPanel")
    maxColWidth = settings.get("maxColWidth")
    loginTimeout = settings.get("loginTimeout")
    queryTimeout = settings.get("queryTimeout")

    showResultsInPanel = showResultsInPanel or False
    optFixedLengthTypeWidth = "-Y {}".format(maxColWidth) if maxColWidth != None else ""
    optVariableLengthTypeWidth = "-y {}".format(maxColWidth) if maxColWidth != None else ""
    optLoginTimeout = "-l {}".format(loginTimeout) if loginTimeout != None else ""
    optQueryTimeout = "-t {}".format(queryTimeout) if queryTimeout != None else ""
    optOutputPath = "-o {}".format(outputPath) if outputPath != None and outputPath != "" else ""
    
    # sqlcmd = "& sqlcmd -E -S {} -d {} -i {} {} {} 2>&1 | Out-File {} -Encoding UTF8".format(server, database, inputPath, optLoginTimeout, optQueryTimeout, outputPath)
    sqlcmd = "& sqlcmd -p -E -S {} -d {} -i {} -f 65001 {} {} {} {} {}".format(server, database, inputPath, optLoginTimeout, optQueryTimeout, optOutputPath, optFixedLengthTypeWidth, optVariableLengthTypeWidth)
    pscmd = "powershell -NoProfile -NonInteractive -NoLogo -Command {}".format(sqlcmd)
    # Create results dir
    create_dir(outputDir)
    # Execute SQL
    CREATE_NO_WINDOW = 0x08000000
    process = subprocess.Popen(pscmd, creationflags = CREATE_NO_WINDOW)
    process.wait()
    if showResultsInPanel:
        results = load_results(outputPath)
        # Create results panel
        isUnlisted = False
        panelName = 'ResultsOfQuery'
        panelFullname = 'output.{}'.format(panelName)
        view = window.create_output_panel(panelName, isUnlisted)
        # Show results
        # AutoIndent is messing with our output if it's not disabled.
        view.settings().set("auto_indent", False)
        view.run_command('insert', {"characters" : results})
        window.run_command('show_panel', {"panel" : panelFullname})
    show_msg("done (elapsed time: {})".format(elapsedTime))

def get_output_dir():
    base_dir = sublime.packages_path()
    rel_dir = "User\\Queries"
    dir = os.path.join(base_dir, rel_dir)
    return dir

def create_dir(dir):
    try:
        os.mkdir(dir)
    except OSError as exc:
        if exc.errno != errno.EEXIST:
            raise
        pass

def load_results(filename):
    # Todo: handle exceptions
    # Ex: opening with the wrong encoding,
    # Ex: the file doesn't even exist (because of previous error with sqlcmd)
    encoding = 'utf-8-sig'
    with open(filename, 'r', encoding=encoding) as f:
        data = f.read()
    return data