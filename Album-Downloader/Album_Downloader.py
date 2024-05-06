import os, glob, win32api, win32con
import tkinter as tk
from mutagen.mp3 import MP3  
from mutagen.easyid3 import EasyID3
from threading import *

#
#------------------------------------------------------DISCLAIMER--------------------------------------------------------------------------------------------------
#
#   1:  This project is purely for eductation and trial purposes
#   2:  Support the Artist wherever and whenever you can via YouTube, Spotify, Bandcamp, or wherever the artist officially hosts or sells their music
#       Examples:
#       -   If you have wifi or unlimited data it is recommended to use those connections to stream the album to support the artist whenever possible, or
#       -   If you like the album and have the disposable income please purchase the music from an official source to listen to it offline
#   3:  If you care about hi-fi versions of the album, THIS IS NOT IT. Please buy the album or listen to it on a different platform
#   4:  This does NOT completely format the album and some manual work NEEDS to be done. This project just eliminates a large portion of the work
#
#   5:  This project uses OS and GLOB libraries. Please read at the code before running it to AVOID running POTENTIALLY MALICIOUS code on your computer
#       -   Never trust an unknown programmer. ALWAYS look over their code before running it, even if a description/summary is provided
#       -   If you still cannot trust the code, make it yourself
#
#------------------------------------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------------------------------------
#
#   The purpose of this program is to be able to download a playlist from YouTube and format it into an album folder
#   1:  Tkinter is used to handle the user input for the save_dir and the YouTube link
#   2:  OS is used to move to the save_dir
#   3:  OS is used to run the YT-DLP command to download the playlist as .mp3 files
#   4:  OS is used to run the YT-DLP command to download the playlist thumbnail as .jpg
#   4:  Mutagen is used to add the remaining metadata to the songs
#   5:  OS is used to move all of the files from the playlist into a folder
#   6:  Win32api and Win32con are used to format the folder as read-only
#
#   Notes:
#   1:  Threading is used to ensure the tkinter window is still responsive while processing the video
#       -   Despite this it is recommended to only download one album at a time to reduce CPU and YouTube server load
#       -   Reducing YouTube server load lessens the likelyhood of being banned or blocked
#   2:  Python does not have the ability to set the folder picture so that has to be done manually afterwards (this can be done in a few seconds)
#       1:  Right-click the album folder
#       2:  Click properties
#       3:  Click customize
#       4:  Optimize for Music, 
#       5:  Check the box to apply template to all subfolders
#       6:  Click Choose File
#       7:  Select image
#       8:  Click open
#       9:  Click apply
#       10: Exit window
#   3:  This project is still a work in progress and may update depending on future changes to YT-DLP
#
#   Future Plans:
#   1:  Add an alias function that keeps a memory, in an external file, of artist names and what the user wants them to be called
#
#   Documentation
#   1: YT-DLP:      https://github.com/yt-dlp/yt-dlp
#   2: Mutagen:     https://mutagen.readthedocs.io/en/latest/index.html
#   3: OS:          https://docs.python.org/3/library/os.html
#   4: GLOB:        https://docs.python.org/3/library/glob.html
#   5: WIN32API:    https://timgolden.me.uk/pywin32-docs/win32api.html
#   6: WIN32CON:    https://timgolden.me.uk/pywin32-docs/win32console.html
#   7: Threading:   https://docs.python.org/3/library/threading.html
#   8: Tkinter:     https://docs.python.org/3/library/tk.html
#
#------------------------------------------------------CODE STARTS HERE-------------------------------------------------------------------------------------------

''' Downloads a Youtube Album from the playlistLink and indexes them by the order they appear in the playlist
    @Params:    path        |   String  |   Path to save the Playlist at
                playlistLink|   String  |   URL to the Playlist to download
    @Returns:   none'''
def downloadAlbum(save_dir, playlistLink):
    os.chdir(f"{save_dir}")
    yt_dlp_command = f'yt-dlp {playlistLink} -f ba -x --audio-format mp3 -o "%(channel)s-+-%(playlist)s-+-%(title)s.%(ext)s" --parse-metadata "playlist_index:%(track_number)s" --embed-metadata'
    os.system(yt_dlp_command)
    thumbnail_command = f'yt-dlp {playlistLink} --skip-download --write-thumbnail -o "thumbnail:"'
    os.system(thumbnail_command)
    structure_album(save_dir)
    print(f"\n\nFinished Downloading and formatting Album to {save_dir}")
    if check_button_2.get() == 1:
        print("\n\nClearing Directory Input")
        directory_path.delete("1.0", "end-1c")
        directory_path.update()
        check_button_2.set(0)
    if check_button_3.get() == 1:
        print("\n\nClearing Playlist Input")
        youtube_url.delete("1.0", "end-1c")
        youtube_url.update()
        check_button_3.set(0)
    if check_button_1.get() == 1:
        check_button_1.set(0)

''' Gets the name of the Album Artist by counting the instances of every unique album artist and sorting them in decreasing order so the most common one is first
    @Params:    album_path      |   String  |   Path to find the Ablum at
    @Returns:   album_artist[0] |   String  |   Name of the Album Artist'''
def get_album_artist(album_path):
    file_list = glob.glob(album_path + "\*.mp3")
    artist_count = {}
    for file in file_list:
        audio = EasyID3(file)
        if 'albumartist' in audio:
            album_artist = audio['albumartist'][0]
            if album_artist in artist_count:
                artist_count[album_artist] += 1
            else:
                artist_count[album_artist] = 1
    sorted_artists = sorted(artist_count.items(), key=lambda x: x[1], reverse=True)
    album_artist = list(sorted_artists)[0]
    return album_artist[0]

''' Sets the Album Artist metadata tag for all of the songs based on the output of get_album_artist()
    @Params:    album_path  |   String  |   Path to find the Album at
                album_artist|   String  |   Name of the Album Artist
    @Returns:   none'''
def set_album_artist(album_path, album_artist):
    file_list = glob.glob(album_path + "\*.mp3")
    for file in file_list:
        song = MP3(file, ID3=EasyID3)
        song['albumartist'] = album_artist
        song.save()

''' Adds meta data to the songs in the album according to their folder and name using Mutagen
    @Params:    current_file    |   String  |   Path to save the metadata to
                channel         |   String  |   Channel/Artist name
                album           |   String  |   Album name
    @Returns:   none'''
def add_tags_to_album(current_file, channel, album):
    song = MP3(current_file, ID3=EasyID3)
    song['album'] = album
    song['albumartist'] = channel
    song.save()

''' Moves the recently downloaded files into their respective folders (creating new ones if needed, and formats them for music albums)
    @Params:    search_path |   String  |   Path to find the album files at
    @Returns:   none'''
def structure_album(search_path):
    print("\n\nFormatting files into Album")
    album_path = ""
    file_list = glob.glob(search_path + "\*.mp3")
    for i in range(len(file_list)):
        current_file_path = file_list[i]
        split_file_path = current_file_path.split("\\")
        current_file_name = split_file_path[len(split_file_path)-1]
        split_file_name = current_file_name.split("-+-")
        channel = split_file_name[0].strip()
        album = split_file_name[1].strip()
        current_file_title = split_file_name[2]

        print("\n\nSwitching to Adding Tags")
        print(f"Adding tags to file {i+1}/{len(file_list)}: {current_file_path}")
        add_tags_to_album(current_file_path, channel, album)

        print("\n\nSwitching to moving files")
        album_path = f"{search_path}\{album}"
        new_file_path = f"{search_path}\{album}\{current_file_title}"
        if os.path.exists(album_path):
            print(f"Moving file {i+1}/{len(file_list)}: {current_file_path} --> {new_file_path}")
            os.rename(current_file_path, new_file_path)
        else:
            print(f"Creating {album_path} folder")
            os.makedirs(album_path)
            print(f"Moving file {i+1}/{len(file_list)}: {current_file_path} --> {new_file_path}")
            os.rename(current_file_path, new_file_path)

    print("\n\nMoving the thumbnail to the correct folder")
    thumbnails = glob.glob(search_path + "\*.jpg")
    thumbnail_path = ""
    for i in range(len(thumbnails)):
        current_thumbnail = thumbnails[i].split(" ")
        thumbnail_path =  " ".join(current_thumbnail[:len(current_thumbnail)-1]) + ".jpg"
        split_thumbnail_path = thumbnail_path.split("\\")
        thumbnail_name = split_thumbnail_path[len(split_thumbnail_path)-1]
        thumbnail_path = f"{album_path}\{thumbnail_name}"
        print(f"Moving file: {thumbnails[i]} --> {thumbnail_path}")
        os.rename(thumbnails[i], thumbnail_path)
        
    print("\n\nSetting the Album File to Read-Only")
    win32api.SetFileAttributes(album_path, win32con.FILE_ATTRIBUTE_READONLY)
    
    print("\n\nChanging the Album Artist")
    set_album_artist(album_path, get_album_artist(album_path))

''' Gets the Path to the Default Directory file regardless of whether it exists or not
    There is a check whether the file exists outside this function scope
    @Params: None
    @Returns: None'''
def get_default_directory():
    current_file_path = os.path.realpath(__file__)
    split_dir = current_file_path.split("\\")
    folder_path = "\\".join(split_dir[:len(split_dir)-1])
    save_dir_file_path = folder_path + "\default save dir.txt"
    return save_dir_file_path

''' Sets the default download path in a text file in the same path as the python file
    There is a check whether save_dir exists outside this function scope
    @Params:    save_dir    |   String  |   Path to save the Album at and the directory to save as default
    @Returns:   none'''
def set_default_download_path(save_dir):
    save_dir_file_path = get_default_directory()
    if os.path.exists(save_dir_file_path):                  # Runs if the text file exists
        print(f"Default Directory File detected. Opening: {save_dir_file_path}")
        print(f"Overwriting Default Directory. New Default: {save_dir}\n\n")
        save_file = open(save_dir_file_path, 'w').close()   # Clears the current text in the file
        save_file = open(save_dir_file_path, 'w')
        save_file.write(save_dir)
        save_file.close()
    else:                                                   # Runs if the text file does not exist
        print(f"Creating Default Directory Text File: {save_dir_file_path}\n\n")
        save_file = open(save_dir_file_path, 'w')
        save_file.write(save_dir)
        save_file.close()

''' Handles the Inputs and Runs the selected program
    This function does not have any Parameters or Returns, but accesses the text from the directory and input_url text boxes
    @Params:    none
    @Returns:   none'''
def run_handler():
    save_dir = directory_path.get("1.0", "end-1c").strip()
    playlist_link = youtube_url.get("1.0", "end-1c").strip()

    run_program = True
    if playlist_link == "" or not playlist_link.startswith("https://www.youtube.com/playlist?list="):
        label_message.config(text="Error: Invalid Playlist URL")
        run_program = False
    if save_dir == "":
        if os.path.exists(get_default_directory()):
            default_dir_file = open(get_default_directory(), 'r')
            save_dir = default_dir_file.read()
            print(f"Using default Directory. Saving files to: {save_dir}\n\n")
        else:
            label_message.config(text="Error: Directory not provided, default not found. Please insert a Directory")
            run_program = False
    if not os.path.exists(save_dir):
        label_message.config(text="Error: Invalid Directory")
        run_program = False

    if run_program:
        if check_button_1.get() == 1:
            label_message.config(text=f"Downloading Album. Setting Default Directory: {save_dir}")
            set_default_download_path(save_dir)
        else:
            label_message.config(text="Downloading Album")
        thread1 = Thread(target=downloadAlbum, args=(save_dir, playlist_link))
        thread1.start()

''' Inserts a directory example into the directory text box
    This function does not have any Parameters or Returns, but modifies the text from the directory text box
    @Params:    none
    @Returns:   none'''
def show_example_directory():
    directory_path.delete("1.0", "end-1c")
    directory_path.insert("1.0", "C:\\Users\\User\\Save File")
    directory_path.update()

''' Inserts a playlist url example into the input_url text box
    This function does not have any Parameters or Returns, but modifies the text from the input_url text box
    @Params:    none
    @Returns:   none'''
def show_example_playlist_url():
    youtube_url.delete("1.0", "end-1c")
    youtube_url.insert("1.0","https://www.youtube.com/playlist?list=PLAYLIST")
    youtube_url.update()

''' Clears all text from the directory text box
    This function does not have any Parameters or Returns, but modifies the text from the directory text box
    @Params:    none
    @Returns:   none'''
def clear_directory():
    directory_path.delete("1.0", "end-1c")
    directory_path.update()

''' Clears all text from the input_url text box
    This function does not have any Parameters or Returns, but modifies the text from the input_url text box
    @Params:    none
    @Returns:   none'''
def clear_youtube():
    youtube_url.delete("1.0", "end-1c")
    youtube_url.update()

''' Clears all text from the directory and input_url text boxes
    This function does not have any Parameters or Returns, but modifies the text from the directory and input_url text boxes
    @Params:    none
    @Returns:   none'''
def clear_input():
    directory_path.delete("1.0", "end-1c")
    youtube_url.delete("1.0", "end-1c")
    directory_path.update()
    youtube_url.update()
    check_button_1.set(0)
    check_button_2.set(0)
    check_button_3.set(0)

root = tk.Tk()
root.geometry("800x350")

# Creates Labels for user convenience
label_directory = tk.Label(text = "Directory")
label_youtube = tk.Label(text = "YouTube URL")
label_options = tk.Label(text = "Options")
label_output = tk.Label(text = "Output")
label_message = tk.Label(text = "")

# Creates Text boxes for user input
youtube_url = tk.Text(root, height = 1, width = 50, bg = "light cyan")
directory_path = tk.Text(root, height = 1, width = 50, bg = "light green")

# Creates Buttons
button_directory = tk.Button(root, text = 'Directory Example', command = lambda: show_example_directory())
button_youtube = tk.Button(root, text = 'Playlist Example', command = lambda: show_example_playlist_url())
button_clear_directory = tk.Button(root, text = 'Clear Directory', command = lambda: clear_directory())
button_clear_youtube = tk.Button(root, text = 'Clear YouTube URL', command = lambda: clear_youtube())
button_clear_all = tk.Button(root, text = 'Clear All Input', command = lambda: clear_input())
button_run = tk.Button(root ,text = "Convert", command = lambda: run_handler())

# Creates Checkbutton
check_button_1 = tk.IntVar()
check_button_2 = tk.IntVar()
check_button_3 = tk.IntVar()
check_set_default = tk.Checkbutton(root, text = "Set as Default Directory", variable = check_button_1, onvalue = 1, offvalue = 0, height = 1, width = 20)
check_clear_directory = tk.Checkbutton(root, text = "Clear Directory after completion", variable = check_button_2, onvalue = 1, offvalue = 0, height = 1, width = 30)
check_clear_youtube = tk.Checkbutton(root, text = "Clear YouTube url after completion", variable = check_button_3, onvalue = 1, offvalue = 0, height = 1, width = 30) 

# Places all of the window features here
label_directory.place(x=20, y=50)
directory_path.place(x=100, y=50)
button_directory.place(x=525, y=50)
button_clear_directory.place(x=650, y=50)

label_youtube.place(x=20, y=100)
youtube_url.place(x=100, y=100)
button_youtube.place(x=525, y=100)
button_clear_youtube.place(x=650, y=100)

label_options.place(x=20, y=150)
check_set_default.place(x=85, y=150)
check_clear_directory.place(x=74, y=175)
check_clear_youtube.place(x=82, y=200)

button_run.place(x=300, y=250)
button_clear_all.place(x=400, y=250)

label_output.place(x=20,y=300)
label_message.place(x=100, y=300)

# Run the window
root.mainloop()