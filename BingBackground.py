import urllib.request
import socket
import datetime
import os, sys
import os.path as op
import json
from PIL import Image
from io import BytesIO
import ctypes
import win32api
import win32con
import win32gui


def delete_old_image(basedir):
    """
    Delete the picture downloaded one week ago.
    """
    print('Starting delete old image...')
    imgdel = 0
    imgfiles = [op.join(basedir, f) for f in os.listdir(basedir) if op.isfile(op.join(basedir, f))]
    for img in imgfiles:
        delta = datetime.date.fromtimestamp(op.getctime(img)) - datetime.date.today()
        if delta.days > 7:
            os.remove(img)
            imgdel += 1
            print('Image %s was deleted...' % (op.basename(img)))
    print('Total %i old image(s) deleted...' % (imgdel))
    return


def download_image(basedir, strProxy=''):
    """
    If has proxy, need setup proxy before open the url
    """
    urlFull = ''
    imgFullpath = ''
    urlDomain = r"http://www.bing.com"
    urlGetJson = r"http://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US"

    # Setup proxy
    if len(strProxy) > 0:
        proxy_handle = urllib.request.ProxyHandler({'http': strProxy})
        opener = urllib.request.build_opener(proxy_handle)
        urllib.request.install_opener(opener)

    # According to bing json string to get the real url of image address
    print('Start getting image address...')
    try:
        html = urllib.request.urlopen(urlGetJson, timeout=3).read()
    except (urllib.error.HTTPError, urllib.error.URLError) as e:
        print('Error: Cannot open the website(%s) because %s, please retry in browser manually...' % (urlGetJson, e))
    except (socket.timeout):
        print('Error: Accessing URL(%s) timeout...' % urlGetJson)
    else:
        print('Downloading JSON string from bing.com...')
        jsonString = json.loads(html)
        print('Configuring JSON string...')
        urlImage = jsonString['images'][0]['url']
        print('Configuring JSON string finished...')
        listTemp = urlImage.split('/')
        imgName = listTemp[len(listTemp) - 1]
        imgFullpath = op.join(basedir, imgName.replace('jpg', 'bmp'))
        urlFull = urlDomain + urlImage
        print('Image web address: (%s)...' % (urlFull))

    # Save the image to local drive
    # **Note here must convert the jpg to bmp at win7 environment**
    if not op.exists(imgFullpath):
        if urlFull != '':
            print('Start downloading image...')
            try:
                imgBinary = urllib.request.urlopen(urlFull, timeout=3).read()
            except (urllib.error.HTTPError, urllib.error.URLError) as e:
                print('Error: Cannot open the website(%s) because %s, please retry in browser manually...' % (urlFull, e))
            except (socket.timeout):
                print('Error: Accessing URL(%s) timeout...' % (urlFull))
            else:
                Image.open(BytesIO(imgBinary)).save(imgFullpath, 'bmp')
                print('Image saved to (%s)...' % (imgFullpath))
                return imgFullpath
        else:
            return ''
    else:
        print('Image today is existing, stop to download...')
        return imgFullpath


def change_background(imgPath):
    key = win32api.RegOpenKey(win32con.HKEY_CURRENT_USER, r'Control Panel\Desktop', 0, win32con.KEY_ALL_ACCESS)
    win32api.RegSetValueEx(key, "WallpaperStyle", 0, win32con.REG_SZ, "0")
    win32api.RegSetValueEx(key, "TileWallpaper", 0, win32con.REG_SZ, "0")
    win32api.RegSetValueEx(key, "PicturePosition", 0, win32con.REG_SZ, "10")
    win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, imgPath, 1 or 2)
    print('Background updated...')
    return

if getattr(sys, 'frozen', False):
    # frozen
    dir_ = os.path.dirname(sys.executable)
else:
    # unfrozen
    dir_ = os.path.dirname(os.path.realpath(__file__))

imgdir = op.join(op.abspath(dir_), 'Wallpapers')
if not op.exists(imgdir):
    os.mkdir(imgdir)

delete_old_image(imgdir)

newImgPath = ''
tryround = 0
while (tryround <= 5 and newImgPath == ''):
    newImgPath = download_image(imgdir, '10.112.254.132:8887')
    if newImgPath == '':
        print('Error: Background download failed, start retry...')
        tryround += 1

if newImgPath != '':
    change_background(newImgPath)
else:
    print('Error: Background update failed, please manually check if the website is available...')
