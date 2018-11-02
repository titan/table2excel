#! /usr/bin/python

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

def table_do_action(action, ctx):
  from table_fsm import Action
  if action == Action.ERROR:
    print("Invalid table format at col %d in line %d" % (ctx.col, ctx.line))
    exit(-1)
  elif action == Action.APPEND:
    ctx.buf += ctx.ch
  elif action == Action.CELL:
    ctx.cells.append(ctx.buf.strip())
    ctx.buf = ''
  elif action == Action.LINE:
    ctx.lines.append(ctx.cells)
    ctx.cells = []
  elif action == Action.ROW:
    cells = []
    for i in range(len(ctx.lines[0])):
      cells.append([])
    for row in range(len(ctx.lines)):
      for col in range(len(ctx.lines[row])):
        if len(ctx.lines[row][col]) > 0:
          cells[col].append(ctx.lines[row][col])
    row = []
    for c in cells:
      row.append('\n'.join(c))
    ctx.rows.append(row)
    ctx.lines = []

def load(src: str):
  from table_fsm import Event, FSM
  ctx = ParsingContext()
  fsm = FSM(table_do_action)
  with open(src, 'r') as input:
    content = input.read(-1)
    for ch in content:
      ctx.ch = ch
      if ch == '\n':
        ctx.line += 1
        ctx.col = 1
        fsm.process(Event.LF, ctx)
      elif ch == '+':
        fsm.process(Event.PLUS, ctx)
        ctx.col += 1
      elif ch == '-':
        fsm.process(Event.MINUS, ctx)
        ctx.col += 1
      elif ch == '|':
        fsm.process(Event.PIPE, ctx)
        ctx.col += 1
      else:
        fsm.process(Event.OTHERS, ctx)
        ctx.col += 1
  return ctx.rows

def save(model, dst: str):
  import xlsxwriter
  wb = xlsxwriter.Workbook(dst)
  ws = wb.add_worksheet()
  for row in range(len(model)):
    for col in range(len(model[row])):
      ws.write(row, col, model[row][col])
  wb.close()

def main(src: str, dst: str):
  model = load(src)
  save(model, dst)

if __name__ == '__main__':
  import argparse
  import sys
  parser = argparse.ArgumentParser()
  parser.add_argument("src", help="The table file")
  parser.add_argument("dst", help="The execel file")
  args = parser.parse_args()
  main(args.src, args.dst)
