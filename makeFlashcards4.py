#!/usr/bin/python
# -*- coding: utf-8 -*-
# makeFlashcards4.py
#
# A Libreoffice-script written in Python to
# create neat flash cards for duplex printing.
#
# Rewritten for Libreoffice 4.0
#
# Stefan Schroeder
# New York, 2013-06-26
#

import uno

HEIGHT = 6
WIDTH = 2


class flashTable:

    """This class creates the flashcards table"""

    def __init__(
        self,
        doc,
        HEIGHT,
        WIDTH,
        complete_list,
        ):
        self.complete_list = complete_list
        self.doc = doc
        self.HEIGHT = HEIGHT
        self.WIDTH = WIDTH
        self.FONTSIZE_BIG = 14.0
        self.FONTSIZE_SMALL = 10.0
        self.BOX_HEIGHT = 3500  # 35mm

    def create_double_page(self, n):
        self.create_table_page(n, 0)
        self.create_table_page(n, 1)

    def create_table_page(self, offset, side):
        """Assemble a table on a page, distinguishing between front side=0
        and back side=1, the offset is the index in complete_list of the first word 
        that appears on this page."""

        if offset >= len(self.complete_list):
            return

        text = self.doc.Text
        table = self.doc.createInstance('com.sun.star.text.TextTable')
        table.initialize(self.HEIGHT, self.WIDTH)

        cursor = text.createTextCursor()
        text.insertTextContent(cursor, table, 0)

        r = table.getRows()
        for i in range(self.HEIGHT):  # with all rows
            r2 = r.getByIndex(i)  # set row properties
            r2.setPropertyValue('Height', self.BOX_HEIGHT)
            r2.setPropertyValue('IsAutoHeight', False)
            for j in range(self.WIDTH):  # with all columns
                coord = str(chr(65 + j)) + str(i + 1)  # first cell is A1, not A0
                if side == 0:
                    item = offset + i * self.WIDTH + j
                    if item > len(self.complete_list) - 1:
                        continue
                    self.insertTextIntoCell(table, coord,
                            self.complete_list[item][0])
                elif side == 1:
                    item = offset + i * self.WIDTH + self.WIDTH - 1 - j
                    if item > len(self.complete_list) - 1:
                        continue

                    # reverse order: cursor doesn't move to end of string

                    if len(self.complete_list[item]) > 2:
                        self.insertTextIntoCell(table, coord, '\n'
                                + self.complete_list[item][2],
                                bigfont=False)
                    self.insertTextIntoCell(table, coord,
                            self.complete_list[item][1], bold=True)

        self.insertTextIntoCell(table, 'A1', '', 0)  # added to solve flushing problems

    def insertTextIntoCell(
        self,
        table,
        cellName,
        mytext,
        bold=False,
        bigfont=True,
        ):
        """Puts some text in a table cell with some basic formatting"""

        tableText = table.getCellByName(cellName)
        cursor = tableText.createTextCursor()
        cursor.setPropertyValue('CharPosture', 0)  # no italic
        cursor.setPropertyValue('CharPostureAsian', 0)  # no italic
        if not bold:
            cursor.setPropertyValue('CharWeight', 0)  # No bold
            cursor.setPropertyValue('CharWeightAsian', 0)  # No bold
        else:
            cursor.setPropertyValue('CharWeight', 150)  # bold
            cursor.setPropertyValue('CharWeightAsian', 150)  # bold
        if bigfont:
            cursor.setPropertyValue('CharHeight', self.FONTSIZE_BIG)  # fontsize
            cursor.setPropertyValue('CharHeightAsian',
                                    self.FONTSIZE_BIG + 4.0)  # fontsizeAsian
        else:
            cursor.setPropertyValue('CharHeight', self.FONTSIZE_SMALL)  # fontsize
            cursor.setPropertyValue('CharHeightAsian',
                                    self.FONTSIZE_SMALL + 4.0)  # fontsizeAsian

        table.Split = False  # Prevent Page Break in Table
        cursor.setPropertyValue('ParaAdjust', 3)  # Center horizontally
        tableText.VertOrient = 2  # Center vertically
        tableText.insertString(cursor, mytext, 0)


class listTable:

    """This class creates a list table with the same data"""

    def __init__(self, doc, complete_list):
        self.complete_list = complete_list
        dummy = ['Front', 'Back', 'Back example']
        self.complete_list.insert(0, dummy)
        self.doc = doc
        self.HEIGHT = len(complete_list)
        self.WIDTH = len(complete_list[1])  # 0 is dummy which is alwa. len()=3

    def create(self):
        text = self.doc.Text
        table = self.doc.createInstance('com.sun.star.text.TextTable')
        table.initialize(self.HEIGHT, self.WIDTH)
        cursor = text.createTextCursor()
        table.setPropertyValue('BreakType', 4)  # PAGE_BEFORE
        text.insertTextContent(cursor, table, 0)

        for i in range(self.HEIGHT):
            for j in range(self.WIDTH):
                tableText = table.getCellByPosition(j, i)
                cursor = tableText.createTextCursor()
                s = self.complete_list[i][j]
                tableText.insertString(cursor, s, 0)
        tableText.insertString(cursor, '', 0)  # added to solve flushing problems


def insertTextIntoCell(
    table,
    cellName,
    text,
    color,
    ):
    tableText = table.getCellByName(cellName)
    cursor = tableText.createTextCursor()
    cursor.setPropertyValue('CharColor', color)
    tableText.setString(text)


def createTable():
    """creates a new writer document and inserts a table with some data"""

    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = \
        smgr.createInstanceWithContext('com.sun.star.frame.Desktop',
            ctx)

    SEPARATOR = '\t'
    model = desktop.getCurrentComponent()
    text = model.Text
    cursor = text.createTextCursor()
    cursor.gotoStart(False)
    cursor.gotoEnd(True)
    myline = text.getString()

    # read separated lines into array

    complete_list = []
    for i in myline.splitlines():
        if SEPARATOR not in i:
            continue
        fields = i.strip().split(SEPARATOR)
        if len(fields) == 2:
            fields.append(' ')
        complete_list.append(fields)

    if len(complete_list) == 0:
        complete_list = [['no', 'words', 'found']]

    # open a writer document

    doc = desktop.loadComponentFromURL('private:factory/swriter',
            '_blank', 0, ())

    text = doc.Text
    cursor = text.createTextCursor()
    table = doc.createInstance('com.sun.star.text.TextTable')

    # Create flash card table

    f = flashTable(doc, HEIGHT, WIDTH, complete_list)
    number_of_pages = int(len(complete_list) / (WIDTH * HEIGHT) + 1)
    for i in range(number_of_pages):
        f.create_double_page(i * WIDTH * HEIGHT)

    # Create list table

    l = listTable(doc, complete_list)
    l.create()


g_exportedScripts = (createTable, )

