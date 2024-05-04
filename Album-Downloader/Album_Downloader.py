import os, glob, win32api, win32con
import tkinter as tk
from moviepy.editor import *
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
#   2:  Implement a default download location. If the save_dir is left blank then it will download the album there. This will also likely be saved externally
#   3:  Implement safety mechanisms for various use-cases that are unaccounted for
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
#   9: Moviepy:     https://zulko.github.io/moviepy/
#
#------------------------------------------------------CODE STARTS HERE-------------------------------------------------------------------------------------------

''' Downloads a Youtube Album from the playlistLink and indexes them by the order they appear in the playlist
    @Params:    path        |   String  |   Path to save the Playlist at
                playlistLink|   String  |   URL to the Playlist to download
    @Returns:   none'''
def downloadAlbum(path, playlistLink):
    os.chdir(f"{path}")
    yt_dlp_command = f'yt-dlp {playlistLink} -f ba -x --audio-format mp3 -o "%(channel)s-+-%(playlist)s-+-%(title)s.%(ext)s" --parse-metadata "playlist_index:%(track_number)s" --embed-metadata'
    os.system(yt_dlp_command)
    thumbnail_command = f'yt-dlp {playlistLink} --skip-download --write-thumbnail -o "thumbnail:"'
    os.system(thumbnail_command)
    structure_album(path)
    print("\n\n\n\nFinished")

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
    print("\n\n\n\nMoving the songs to the correct folder")
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

        add_tags_to_album(current_file_path, channel, album)

        album_path = f"{search_path}\{album}"
        new_file_path = f"{search_path}\{album}\{current_file_title}"
        if os.path.exists(album_path):
            print(f"Moving file {i+1}/{len(file_list)}: {current_file_path} --> {new_file_path}")
            os.rename(current_file_path, new_file_path)
        else:
            os.makedirs(album_path)
            print(f"Moving file {i+1}/{len(file_list)}: {current_file_path} --> {new_file_path}")
            os.rename(current_file_path, new_file_path)

    print("\n\n\n\nMoving the thumbnail to the correct folder")
    thumbnails = glob.glob(search_path + "\*.jpg")
    thumbnail_path = ""
    for i in range(len(thumbnails)):
        #thumbnail_path = thumbnails[i][:len(thumbnails[i])-41] + ".jpg"         # try current_thumbnail = thumbnails[i].split(" ") and thumbnailpath =  current_thumbnail[:len(current_thumbnail)-1].join(" ") + ".jpg"
        current_thumbnail = thumbnails[i].split(" ")
        thumbnail_path =  " ".join(current_thumbnail[:len(current_thumbnail)-1]) + ".jpg"
        split_thumbnail_path = thumbnail_path.split("\\")
        thumbnail_name = split_thumbnail_path[len(split_thumbnail_path)-1]
        thumbnail_path = f"{album_path}\{thumbnail_name}"
        print(f"Moving file: {thumbnails[i]} --> {thumbnail_path}")
        os.rename(thumbnails[i], thumbnail_path)
        
    print("\n\n\n\nSetting the Album File to Read-Only")
    win32api.SetFileAttributes(album_path, win32con.FILE_ATTRIBUTE_READONLY)
    
    print("\n\n\n\nChanging the Album Artist")
    set_album_artist(album_path, get_album_artist(album_path))
    
''' Handles the Inputs and Runs the selected program
    This function does not have any Parameters or Returns, but accesses the text from the directory and input_url text boxes
    @Params:    none
    @Returns:   none'''
def run_handler():
    save_dir = directory.get("1.0", "end-1c")
    youtube_url = input_url.get("1.0", "end-1c")
    if save_dir != "" and youtube_url != "":
        l4.config(text="Downloading Album")
        thread1 = Thread(target=downloadAlbum, args=(save_dir, youtube_url))
        thread1.start()
    else:
        if save_dir == "":
            l4.config(text="Error: Input not detected: Directory")
        else:
            l4.config(text="Error: Input not detected: Input URL")

''' Inserts a directory example into the directory text box
    This function does not have any Parameters or Returns, but modifies the text from the directory text box
    @Params:    none
    @Returns:   none'''
def show_example_directory():
    directory.delete("1.0", "end-1c")
    directory.insert("1.0", "C:\\Users\\User\\Save File")
    directory.update()

''' Inserts a playlist url example into the input_url text box
    This function does not have any Parameters or Returns, but modifies the text from the input_url text box
    @Params:    none
    @Returns:   none'''
def show_example_playlist_url():
    input_url.delete("1.0", "end-1c")
    input_url.insert("1.0","https://www.youtube.com/playlist?list=PLAYLIST")
    input_url.update()

''' Clears all text from the directory and input_url text boxes
    This function does not have any Parameters or Returns, but modifies the text from the directory and input_url text boxes
    @Params:    none
    @Returns:   none'''
def clear_input():
    directory.delete("1.0", "end-1c")
    input_url.delete("1.0", "end-1c")
    directory.update()
    input_url.update()

root = tk.Tk()
root.geometry("750x275")

# Creates Labels for user convenience
l1 = tk.Label(text = "Directory")
l2 = tk.Label(text = "YouTube URL")
l3 = tk.Label(text = "Output")
l4 = tk.Label(text = "")

# Creates Text boxes for user input
input_url = tk.Text(root, height = 1, width = 50, bg = "light cyan")
directory = tk.Text(root, height = 1, width = 50, bg = "light green")

# Creates Buttons
button1 = tk.Button(root ,text = "Convert", command = lambda: run_handler())
button2 = tk.Button(root, text = 'Directory Example', command = lambda: show_example_directory())
button3 = tk.Button(root, text = 'Playlist Example', command = lambda: show_example_playlist_url())
button4 = tk.Button(root, text = 'Clear Input', command = lambda: clear_input())

# Places all of the window features here
l1.place(x=20, y=50)
directory.place(x=100, y=50)
button2.place(x=525, y=50)

l2.place(x=20, y=100)
input_url.place(x=100, y=100)
button3.place(x=525, y=100)

button1.place(x=300, y=150)
button4.place(x=400, y=150)

l3.place(x=20,y=200)
l4.place(x=100, y=200)

# Run the window
root.mainloop()