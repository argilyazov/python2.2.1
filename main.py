from jinja2 import Environment, FileSystemLoader
import pdfkit

options = {
    "enable-local-file-access": ""
}
config = pdfkit.configuration(wkhtmltopdf=r'D:\AnacondaPy\wkhtmltopdf\bin\wkhtmltopdf.exe')

env = Environment(loader=FileSystemLoader('.'))
template = env.get_template("ez.html")
header = [{f'key': f'key{i}','dfd':'555'} for i in range(10)]
data = [{f'value':i+10,'dfd':'555'} for i in range(10)]

pdf_template = template.render(header)
pdf_template = template.render(data)

pdfkit.from_string(pdf_template, 'ez.pdf', configuration=config, options=options)
