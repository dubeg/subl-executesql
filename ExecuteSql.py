import sublime
import sublime_plugin
import os
import sys
import subprocess
import errno

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
    server = settings.get("server")
    database = settings.get("database")
    loginTimeout = settings.get("loginTimeout")
    optLoginTimeout = "-l {}".format(loginTimeout) if loginTimeout != None else ""
    queryTimeout = settings.get("queryTimeout")
    optQueryTimeout = "-t {}".format(queryTimeout) if queryTimeout != None else ""
    inputPath = filename
    outputDir = get_output_dir()
    outputPath = os.path.join(outputDir, os.path.basename(filename) + ".results")
    sqlcmd = "& sqlcmd -E -S {} -d {} -i {} {} {} 2>&1 | Out-File {} -Encoding UTF8".format(server, database, inputPath, optLoginTimeout, optQueryTimeout, outputPath)
    print(sqlcmd)
    pscmd = "powershell -NoProfile -NonInteractive -NoLogo -Command {}".format(sqlcmd)

    create_dir(outputDir)
    # Execute SQL
    CREATE_NO_WINDOW = 0x08000000
    process = subprocess.Popen(pscmd, creationflags = CREATE_NO_WINDOW)
    process.wait()
    results = load_results(outputPath)
    # Create results panel
    isUnlisted = False
    panelName = 'ResultsOfQuery'
    panelFullname = 'output.{}'.format(panelName)
    view = window.create_output_panel(panelName, isUnlisted)
    # Show results
    view.run_command('insert', {"characters" : results})
    window.run_command('show_panel', {"panel" : panelFullname})
    show_msg("done.")

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
    bytes = min(32, os.path.getsize(filename))
    raw = open(filename, 'rb').read(bytes)
    if raw.startswith(codecs.BOM_UTF8):
        encoding = 'utf-8-sig'
        infile = io.open(filename, 'r', encoding=encoding)
        data = infile.read()
        infile.close()
        return data
    # else:
    #     result = chardet.detect(raw)
    #     encoding = result['encoding']
    return "unexpected error - powershell did not generate the output file."