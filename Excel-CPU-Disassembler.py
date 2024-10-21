import sys
from datetime import datetime
import argparse
import openpyxl


HEADER_STR = '''\
; Source Generated with Excel-CPU-Disassembler.py
; https://github.com/Lil-Ran/Excel-CPU-Disassembler

; No guarantee of accuracy or functionality is provided.
; It is recommended to set tab width to 8 spaces.

; File: {}
; Time: {}

'''


class Address:
    def __init__(self, address):
        self.address = address
    
    def to_hex(self):
        return f'{self.address:04X}'
    
    def to_row_col(self):
        return (self.address // 0x100, self.address % 0x100)
    
    def to_excel(self):
        row, col = self.to_row_col()
        if col >= 26:
            return f"{chr(col // 26 + 65)}{chr(col % 26 + 65)}{row + 1}"
        else:
            return f"{chr(col + 65)}{row + 1}"
    
    @classmethod
    def from_hex(cls, hex_str):
        return cls(int(hex_str, 16))
    
    @classmethod
    def from_row_col(cls, row, col):
        return cls(row * 0x100 + col)
    
    @classmethod
    def from_excel(cls, excel_str):
        if excel_str[1].isalpha():
            col_h = ord(excel_str[0].upper()) - 65
            col_l = ord(excel_str[1].upper()) - 65
            col = col_h * 26 + col_l
            row = int(excel_str[2:]) - 1
        else:
            col = ord(excel_str[0].upper()) - 65
            row = int(excel_str[1:]) - 1
        return cls.from_row_col(row, col)

    def __str__(self):
        global args
        if args.address_style == 'deci':
            return str(self.address)
        elif args.address_style == 'hex':
            return self.to_hex()
        elif args.address_style == 'rowcol':
            return '_'.join(list(map(str, list(self.to_row_col()))))
        elif args.address_style == 'excel':
            return self.to_excel()
    
    def __repr__(self):    # for XREF output
        return str(self)

    def __format__(self, format_spec: str) -> str:
        return str(self).__format__(format_spec)

    def tab_pad(self, width=4):
        return str(self) + ('\t' if len(str(self)) < width else '')


class Cell:
    def __init__(self, data=0):
        self.data = data
        self.label = None  # 'Entrypoint_'; preserve for future
        self.exec_from_prev = None
        self.is_2nd_word = False  # only set to True, never set to False
        self.read_from = []
        self.write_from = []
        self.jump_from = []  # include jmp (no call or ret till now)
    

def preparation():
    global cells, entry, last

    opcode = lambda x: x >> 8
    
    # get entry point
    if cells[0].data == 0:
        entry = cells[1].data
        cells[cells[1].data].label = f'Entrypoint_{Address(cells[1].data)}'
        cells[cells[1].data].jump_from.append(Address(0))
    else:
        entry = 0
    pc = entry

    # find last instruction
    last = 0xffff
    while cells[last].data == 0:
        last -= 1

    # set cell properties
    def one_instruction(pc_):
        global cells
        pc_here = pc_
        if opcode(cells[pc_].data) < 4:  # JXX
            cells[cells[pc_+1].data].jump_from.append(Address(pc_))
        elif opcode(cells[pc_].data) == 4:  # LOAD
            cells[cells[pc_+1].data].read_from.append(Address(pc_))
        elif opcode(cells[pc_].data) == 6:  # STORE
            cells[cells[pc_+1].data].write_from.append(Address(pc_))
        if opcode(cells[pc_].data) < 7:
            cells[pc_+1].is_2nd_word = True
            pc_ += 2
        else:
            pc_ += 1
        if cells[pc_here].data != 0:  # JUMP_ABSOLUTE
            cells[pc_].exec_from_prev = Address(pc_here)
        return pc_
    
    # traverse all cells
    while pc <= last:
        pc = one_instruction(pc)
    
    # cells before entry point
    x = 0
    while x < entry:
        if len(cells[x].jump_from) == 0 and cells[x].exec_from_prev is None:
            x += 1
            continue
        x = one_instruction(x)
    
    # set label
    pc = entry
    while pc <= last:
        if (not cells[pc].is_2nd_word) and (cells[pc].label is None) and (len(cells[pc].jump_from) != 0):
            cells[pc].label = f'loc_{Address(pc).tab_pad()}'
        pc += 1


def output():
    global args, cells, entry, last
    fout = args.output
    if fout != sys.stdout:
        fout.write(HEADER_STR.format(args.input, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    
    opcode = lambda x: x >> 8
    reg1 = lambda x: (x >> 4) & 0xf
    reg2 = lambda x: x & 0xf

    ins_set = ['JMP', 'JEQ', 'JLT', 'JGE', 'LOAD', 'LOAD', 'STORE', 'STORE', 'TRAN',
               'ADD', 'SUB', 'MULT', 'DIV', 'INC', 'DEC', 'AND', 'OR', 'XOR',
               'NOT', 'ROL', 'ROR', 'CMP', 'CLC', 'STC', 'NOP', 'LOAD']
    
    # data section
    if entry != 0:
        # ignore cells[0] and cells[1]
        if not args.no_warnings:
            for i in [0, 1]:
                if len(cells[i].jump_from) != 0: fout.write(f'; Warning: address {Address(i)} may be a jump target of {cells[i].jump_from}\n')
                if len(cells[i].read_from) != 0: fout.write(f'; Warning: address {Address(i)} may be read by {cells[i].read_from}\n')
                if len(cells[i].write_from) != 0: fout.write(f'; Warning: address {Address(i)} may be overwritten by {cells[i].write_from}\n')
        fout.write('\n.DATA\n')
        for i in range(2, entry):
            if len(cells[i].jump_from) != 0 and not args.no_warnings:
                fout.write(f'\t; Warning: the data below (@{Address(i)}) may be a jump target of {cells[i].jump_from}\n')
            fout.write(f'\tvar_{Address(i)} = ${cells[i].data:04X}')
            # if args.include_address: fout.write(f'\t; at {Address(i)}')
            if args.decode_all and opcode(cells[i].data) < 26:
                fout.write(f'\t\t; {ins_set[opcode(cells[i].data)]}')
            fout.write('\n')
            if len(cells[i].read_from) != 0: fout.write(f'\t\t\t\t\t; XREF(R): {cells[i].read_from}\n')
            if len(cells[i].write_from) != 0: fout.write(f'\t\t\t\t\t; XREF(W): {cells[i].write_from}\n')
        fout.write('\n.CODE\n')

    pc = entry
    while pc <= last:
        ins = cells[pc].data
        opcode_ = opcode(ins)
        reg1_ = reg1(ins)
        reg2_ = reg2(ins)
        second = cells[pc+1].data

        unused_bits = 0   # 0: unused bits are all 0;
                          # 1: type of instruction has 4 unused bits, but not all 0 in instance;
                          # 2: type of instruction has 8 unused bits, but not all 0 in instance.

        # up msg
        if cells[pc].label is not None:
            fout.write(f'\n{cells[pc].label}:\n')
        if len(cells[pc].read_from) != 0 and not args.no_warnings:
            fout.write(f'\t; Warning: the code below (@{Address(pc)}) may be read by {cells[pc].read_from}\n')
        if len(cells[pc].write_from) != 0 and not args.no_warnings:
            fout.write(f'\t; Warning: the code below (@{Address(pc)}) may be overwritten by {cells[pc].write_from}\n')
        
        # opcode and operands
        if opcode_ < 4:
            fout.write(f'\t{ins_set[opcode_]}\t')
            if cells[second].label is not None:
                fout.write(f'{cells[second].label}\t')
            else:
                fout.write(f'@{Address(second).to_hex()}\t\t')
            if reg1_ != 0 or reg2_ != 0: unused_bits = 2
        elif opcode_ in [4, 6]:
            tmp_operand_2 = f"var_{Address(second).tab_pad()}" if second<entry else f"@{Address(second).to_hex()}\t"
            fout.write(f'\t{ins_set[opcode_]}\tR{reg1_}\t{tmp_operand_2}')
            if reg2_ != 0: unused_bits = 1
        elif opcode_ == 5:
            fout.write(f'\t{ins_set[opcode_]}\tR{reg1_}\t${second:04X}\t')
            if reg2_ != 0: unused_bits = 1
        elif 7<=opcode_<=12 or 15<=opcode_<=17 or opcode_ in [21, 25]:
            fout.write(f'\t{ins_set[opcode_]}\tR{reg1_}\tR{reg2_}\t')
        elif opcode_ in [13, 14, 18]:
            fout.write(f'\t{ins_set[opcode_]}\tR{reg1_}\t\t')
            if reg2_ != 0: unused_bits = 1
        elif opcode_ in [19, 20]:
            fout.write(f'\t{ins_set[opcode_]}\tR{reg1_}\t#{reg2_}\t')
        elif 22<=opcode_<=24:
            fout.write(f'\t{ins_set[opcode_]}\t\t\t')
            if reg1_ != 0 or reg2_ != 0: unused_bits = 2
        else:
            fout.write(f'\tNOP\t\t\t')    # opcode >= 26

        # include address or data
        if args.include_address: fout.write(f'\t; at {Address(pc)}')
        if args.include_data: fout.write(f'\t; data: {ins:04X}{" %04X"%(cells[pc+1].data) if opcode_<7 else ""}')
        fout.write('\n')

        # opcode >= 26 (down)
        if opcode_ >= 26:
            fout.write(f'\t; Error: instruction above is not NOP but unknown opcode {ins:04X};\n\t  plain data in code section is not allowed by compiler;\n\t  if it is a junk code that would never be read or executed,\n\t  please set Cell {Address(pc).to_excel()} to 6144 (NOP) in Excel manually.\n\n')

        # 2nd word has jump_from
        if opcode_ < 7 and len(cells[pc+1].jump_from) != 0 and not args.no_warnings:
            fout.write(f'\t\t\t\t\t; Warning: the operand of {ins_set(opcode_)} (@{Address(pc+1)}) may be a jump target of {cells[pc+1].jump_from}\n')
        
        # unused bits
        if unused_bits and not args.no_warnings:
            fout.write(f'\t\t\t\t\t; Warning: instruction above has unused bits ({"%01X"%reg1_ if unused_bits==2 else ""}{"%01X"%reg2_})\n')
        
        # XREF and NO XREF warn
        if len(cells[pc].jump_from) != 0:
            fout.write(f'\t\t\t\t\t; XREF(X): {cells[pc].jump_from}\n')
        elif cells[pc].exec_from_prev is None and not args.no_warnings:
            fout.write(f'\t\t\t\t\t; Warning: the code above (@{Address(pc)}) may not be executed\n')

        # decode all - 2nd word
        if args.decode_all and opcode_ < 7 and 0 < opcode(cells[pc+1].data) < 26:    # no 00 - JMP
            fout.write(f'\t\t\t\t\t; -A: operand can be decoded as {ins_set[opcode(cells[pc+1].data)]}\n')
        
        pc += 1 if opcode_>=7 else 2    # include opcode >= 26


def load_excel():
    global args, cells
    cells = [Cell() for _ in range(0x10000)]
    wb = openpyxl.load_workbook(args.input, read_only=True, data_only=True)
    ws = wb.active
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                cells[(cell.row-1) * 0x100 + (cell.column-1)].data = cell.value


def arg_parse():
    global args
    parser = argparse.ArgumentParser(description='An open source Excel CPU disassembler.')
    parser.add_argument('input', type=str, help='input file, must be ROM.xlsx')
    parser.add_argument('-o', '--output', help='output file, default is stdout', default=sys.stdout, type=argparse.FileType('w'))
    parser.add_argument('-s', '--address-style', choices=['deci', 'hex', 'rowcol', 'excel'], help='address style in output file (including lables), default is hex', default='hex')
    parser.add_argument('-a', '--include-address', action="store_true", help='include address of each line in comments')
    parser.add_argument('-d', '--include-data', action="store_true", help='include data of each line in comments')
    parser.add_argument('-A', '--decode-all', action="store_true", help='decode all data (.DATA section and operand of two-word instrutions) as instruction in comments')
    parser.add_argument('-n', '--no-warnings', action="store_true", help='do not output warnings')
    args = parser.parse_args()


if __name__ == '__main__':
    arg_parse()
    load_excel()
    preparation()
    output()
