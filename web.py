

import pdfkit
from default import CSS, STATIC_COLS
import os

def generate_html(df):
    cols = df.columns.values
    columns = '<tr style="page-break-inside: avoid" class="p-3 mb-2 bg-secondary text-white">' + '<th scope="col">#</th>' + ' '.join(['<th scope="col">' + str(c) + '</th>' for c in cols]) + '</tr>'
    data = ''
    for idx in df.index:
        
        data += "<tr>"
        data += f'<th scope="row">{idx+1}</th>'
        for c_i, c in enumerate(cols):
            value = df.at[idx, c]
            color = ''

            if c in STATIC_COLS:
                data += f'<td>' + str(value) + '</td>'
            else:
                if value == 1:
                    color = "p-3 mb-2 bg-danger text-white"
                    # color = "table-danger"
                elif value == 0:
                    color = "p-3 mb-2 bg-success text-white"
                    # color = "table-success"
                elif value == 2:
                    color = "p-3 mb-2 bg-warning text-dark"
                    
                    # color = "table-success"
                if c in ['M12. Интервью натижалари:', 'Ташрифлар сони (шу пайтгача)']:
                    color = "p-3 mb-2 bg-info text-white"
                    # color = "table-info"
                    data += f'<td class="{color}">' + '<b>' + str(value) + '</b>' + '</td>'
                else:
                    data += f'<td class="{color}">' + f"{value}" + '</td>'
                  
                
            
        data += "</tr> "

    html_string = f'''
    <html>
        <head><title>MDP Monitoring Status</title>
            <meta charset="utf-8">
            <meta name="pdfkit-orientation" content="Landscape"/>
            <link href="/css/style.css" rel="stylesheet type="text/css" />
            <link href="/css/bootstrap.css" rel="stylesheet type="text/css" />
            <link href="/css/bootstrap.css" rel="stylesheet type="text/css" />
        </head>
        <body>
            <h2 class="text-center">MDP мониторинг статуси</h2>
            <table class="table table-bordered">
                <thead style="display: table-header-group">
                    {columns}
                </thead>
                <tbody>
                    {data}
                </tbody>
            </table>
        </body>
    </html>
    '''

    # with open('asdasd.html', 'w', encoding='utf-8') as f:
    #     f.write(html_string)
  
    return html_string

def to_pdf(df, pdf_filename):
    html_string = generate_html(df)
    folder = 'pdf'
    if not os.path.isdir(folder):
        os.mkdir(folder)
    filepath = os.path.join(folder, pdf_filename)
    pdfkit.from_string(html_string , filepath, css=CSS) 
    print(filepath)
    return filepath