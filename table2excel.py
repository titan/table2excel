#! /usr/bin/python
from table_fsm import Delegate, StateMachine

class ParsingContext:
  def __init__(self):
    self.buf = ""
    self.tmp = ""
    self.ch = None
    self.line = 1
    self.col = 1
    self.cells = []
    self.lines = []
    self.rows = []
    self.cellwidths = []
    self.cellwidth = 0
    self.cellindex = 0

class ParsingDelegate(Delegate):
  def error(self, ctx):
    print("Invalid table format at col %d in line %d" % (ctx.col, ctx.line))
    exit(-1)
  def reset_cell_index(self, ctx):
    ctx.cellindex = 0
  def append_width(self, ctx):
    ctx.cellwidths.append(ctx.cellwidth)
  def reset_width(self, ctx):
    ctx.cellwidth = 0
  def incr_width(self, ctx):
    ctx.cellwidth += 1
  def append(self, ctx):
    ctx.buf += ctx.ch
  def cell(self, ctx):
    ctx.cells.append(ctx.buf.strip())
    ctx.buf = ''
  def incr_cell_index(self, ctx):
    ctx.cellindex += 1
  def line(self, ctx):
    ctx.lines.append(ctx.cells)
    ctx.cells = []
  def row(self, ctx):
    cells = []
    for i in range(len(ctx.lines[0])):
      cells.append([])
    for row in range(len(ctx.lines)):
      for col in range(len(ctx.lines[row])):
        if len(ctx.lines[row][col]) > 0:
            cells[col].append(ctx.lines[row][col])
    row = []
    for cell in cells:
      if len(cell) > 0:
        content = '\n'.join(cell)
        if content.startswith('----') or content.startswith('===='):
          content = '\n' + content
        if content.endswith('----') or content.endswith('===='):
          content += '\n'
        row.append(content)
      else:
        row.append(None)
    ctx.rows.append(row)
    ctx.lines = []
  def clear_widths(self, ctx):
    ctx.cellwidths = []

def load(src: str):
  ctx = ParsingContext()
  fsm = StateMachine(ParsingDelegate())
  with open(src, 'r') as input:
    content = input.read(-1)
    for ch in content:
      ctx.ch = ch
      if ch == '\n':
        ctx.line += 1
        ctx.col = 1
        fsm.lf(ctx)
      elif ch == '+':
        fsm.plus(ctx)
        ctx.col += 1
      elif ch == '-':
        fsm.minus(ctx)
        ctx.col += 1
      elif ch == '|':
        ctx.col
        splitors = [1]
        for cellwidth in ctx.cellwidths:
          splitors.append(splitors[-1] + cellwidth + 1)
        if ctx.col in splitors:
          fsm.pipe(ctx)
        else:
          fsm.pipe_in_cell(ctx)
        ctx.col += 1
      else:
        fsm.others(ctx)
        ctx.col += 1
  return ctx.rows

def save(model, dst: str):
  from openpyxl import Workbook
  wb = Workbook(write_only = True)
  ws = wb.create_sheet()
  for rid in range(len(model)):
    row = []
    for cid in range(len(model[rid])):
      if model[rid][cid]:
        row.append(model[rid][cid])
      else:
        row.append(None)
    ws.append(row)
  wb.save(dst)

def main(src: str, dst: str):
  model = load(src)
  save(model, dst)

if __name__ == '__main__':
  import argparse
  import sys
  import os.path
  parser = argparse.ArgumentParser()
  parser.add_argument("src", help="The table file")
  parser.add_argument("-d", "--dst", help="The excel file", default = None)
  args = parser.parse_args()
  dst = args.dst
  if dst == None:
    path = os.path.dirname(args.src)
    name = os.path.basename(args.src)
    (n, e) = os.path.splitext(name)
    dst = os.path.join(path, n + ".xlsx")
  main(args.src, dst)
