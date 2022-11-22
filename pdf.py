from jinja2 import Environment, FileSystemLoader
import pdfkit

name = 'Александр'

env = Environment(loader=FileSystemLoader('.'))
template = env.get_template("pdf_template.html")

pdf_template = template.render({'items': items})
pdfkit.from_string(pdf_template, 'out.pdf')