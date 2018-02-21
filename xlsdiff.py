#!/usr/bin/env python
import argparse
import os
import sys
import textwrap
import traceback
import xlrd
import xlwt

class Main:

    def __init__(self, options):
        self.options = options

        self.wb1 = xlrd.open_workbook( options.file1 )
        self.wb2 = xlrd.open_workbook( options.file2 )

        self.wb1data = {}

        self.outwbname = options.outfile
        self.outwb = None

        self.writeidx = 0


    def write_row(self, rowdata):
        for i,col in enumerate(rowdata):
            style = col['style']
            if style is not None:
                self.outwb.write(self.writeidx, i, col['value'], style=style)
            else:
                self.outwb.write(self.writeidx, i, col['value'])

        self.writeidx += 1


    def walk_new(self, ws):
        wb1datakeys = self.wb1data.keys()

        for irow in range(ws.nrows):
            row = ws.row(irow)
            label = row[0].value
            outrow = [{'value': label, 'style': None}]
            oldrow = self.wb1data.get(label, None)

            # Both files have matching row label
            if oldrow:
                wb1datakeys.remove(label)
                newrow = [x.value for x in row[1:]]

                for old,new in zip(oldrow, newrow):
                    # Yellow
                    if old != new:
                        outrow.append({ 'value': new, 'style': xlwt.easyxf('pattern: pattern solid, fore_colour yellow;') })

                    # White
                    else:
                        outrow.append({ 'value': new, 'style': None })

            # This is a new label
            # Green
            else:
                outrow[0]['style'] = xlwt.easyxf( 'pattern: pattern solid, fore_colour green' )
                outrow.extend( [{ 'value': x.value, 'style': xlwt.easyxf( 'pattern: pattern solid, fore_colour green' ) } for x in row[1:]] )


            self.write_row(outrow)

        # get rows that were deleted in the new file
        for key in wb1datakeys:
            outrow = [{ 'value': x, 'style': xlwt.easyxf( 'pattern: pattern solid, fore_colour red' ) } for x in self.wb1data[key]]
            self.write_row(outrow)


    def original_data(self, ws):
        for irow in range(ws.nrows):
            row = ws.row(irow)
            self.wb1data[row[0].value] = [x.value for x in row[1:]]

    def main(self):
        wb = xlwt.Workbook()
        self.outwb = wb.add_sheet('xlate')
        self.original_data(self.wb1.sheet_by_index( 0 ))
        self.walk_new(self.wb2.sheet_by_index( 0 ))

        wb.save(self.outwbname)
        return True


def options():
    parser = argparse.ArgumentParser(
        description=textwrap.dedent( '''
            Shows difference between 2 xls documents (only works on first sheet of both files)
        ''' ),
        add_help=False,
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument( '--help',  help=argparse.SUPPRESS, action='help' )
    parser.add_argument( 'file1',   help='Old File', type=str )
    parser.add_argument( 'file2',   help='New File', type=str )
    parser.add_argument( 'outfile', help='Compared File', type=str )

    return parser.parse_args( )

if __name__ == '__main__':
    try:
        main = Main( options() )
        sys.exit( main.main() )
    except KeyboardInterrupt as e:
        raise e
    except SystemExit as e:
        raise e
    except Exception as e:
        print('ERROR, UNEXPECTED EXCEPTION')
        print(str(e))
        traceback.print_exc()
        os._exit(1)
