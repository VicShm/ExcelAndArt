# Excel And Art
## Image to Excel converter

Some time ago I saw some works by a Japanese artist Tatsuo Horiuchi who decided to use Microsoft Excel as his drawing tool.

I liked it a lot, it's true art!

You can see it [here](https://pasokonga.com/).
Beautiful, isn't it?

Matsuo Horiuchi is an artist, yeah, but I'm a programmer! So how can I use Excel the same way?

Let's get to work!
* Python 3.9.5.
* Python Libraries: xlsxwriter, PIL, time.
* Microsoft Exel 2016, WPS Office 11.2 also work well.

I wrote this script: excelandart.py

Specify image file in the 7th line:
- img_file = './cat/cat.png'

In the 8th line specify the Excel file into wich we convert the image:
- exl_file = './cat/cat.xlsx'

Let's run:
python excelandart.py

And also keep in mind computer performance.
For example my laptop (2012, Intel i5, 4 Gb RAM, SSD) converted cat.png for 57 minutes and same operation took 11 minutes on Intel i7-8700.

And also Excel suddenly failed with an error of the cell format opening the workbook. I  didn't look any further and just сonverted the image from png to jpg and it worked.

Examples are in catalogues cat, eye, girl.

Best Regards.

Victor Shm.

vicshm@gmail.com
