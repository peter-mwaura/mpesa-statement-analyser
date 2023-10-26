from flask import Flask, render_template, url_for, request, make_response
from utils import extract_from_pdf, parse_mpesa_content, find_name, summary, paidin, withdrawal, listing, dfs_tabs
app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def hello_world():
    error = ''
    if request.method == 'POST':
        try:
            numpages, txtfile = extract_from_pdf(request.files['file'], request.form.get('password'))
        except Exception as e:
            error = 'Check file input and password!'
            #error = e
        else:
            title = "unknown"
            content, matches2 = parse_mpesa_content(txtfile)
            if matches2:
                #name = find_name(matches2)
                title = matches2 + ' ' +'MPESA'
            if content:
                #pandas operation
                summary1 = summary(content)
                positives = paidin(content)
                negatives = withdrawal(content)
                dfslist = listing(summary1, positives, negatives)
                sheets = ['SUMMARY', 'PAID IN DATA', 'WITHDRAWN DATA']
                #output = dfs_tabs(content, sheets, content)
                output = dfs_tabs(dfslist, sheets, content)
                resp = make_response((output.getvalue(), {
                    'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    'Content-Disposition': 'attachment; filename={}.xlsx'.format(title)
                }))
            else:
                resp =  make_response({'response' : content}, {
                    'Content-Type': 'application/json',
                })
            return resp

    return render_template('index.html', error=error)

#To remove
if __name__ == '__main__':
    app.run(debug=True)