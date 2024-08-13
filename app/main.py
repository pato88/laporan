from flask import Flask, render_template, send_file, request
import mysql.connector
import pandas as pd
import io

app = Flask(__name__)

def get_data(start_date, end_date):
    connection = mysql.connector.connect(
        host='localhost',
        user='root',
        password='naufal',
        database='gatrans'
    )
    cursor = connection.cursor(dictionary=True)
    query = """
    SELECT 
        mstcustomer.cstPT AS AgentName,
        CASE 
            WHEN SUBSTRING(csFlight, 1, 2) IN ('ID', 'IU', 'IW', 'JT') THEN 'Lion Grup'
            ELSE SUBSTRING(csFlight, 1, 3)  -- Menggunakan 3 karakter untuk flight selain Lion Grup
        END AS FlightGroup,
        SUM(csBeratChargeAble) AS TotalBeratChargeAble
    FROM mstcsc
    JOIN mstcustomer ON mstcsc.csCst = mstcustomer.cstKode
    WHERE csTgl BETWEEN %s AND %s
    GROUP BY AgentName, FlightGroup
    """
    cursor.execute(query, (start_date, end_date))
    data = cursor.fetchall()
    cursor.close()
    connection.close()
    return data

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate_report', methods=['POST'])
def generate_report():
    start_date = request.form.get('start_date')
    end_date = request.form.get('end_date')
    return render_template('index.html', start_date=start_date, end_date=end_date)

@app.route('/download_excel')
def download_excel():
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    data = get_data(start_date, end_date)
    df = pd.DataFrame(data)
    
    # Pivoting the data
    pivot_df = df.pivot(index='AgentName', columns='FlightGroup', values='TotalBeratChargeAble').fillna(0)
    
    # Reset index to make 'AgentName' a column again
    pivot_df.reset_index(inplace=True)

    # Add Total Berat Charge Able column
    pivot_df['Total Berat Charge Able'] = pivot_df.iloc[:, 1:].sum(axis=1)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        pivot_df.to_excel(writer, sheet_name='Report', index=False)
        worksheet = writer.sheets['Report']

        # Formatting the Excel sheet
        worksheet.set_column('A:A', 20)  # Agent Name column
        worksheet.set_column(1, len(pivot_df.columns) - 1, 20)  # Flight columns
        worksheet.write('A1', 'Agent Name')

    output.seek(0)
    return send_file(output, as_attachment=True, download_name='report.xlsx')

if __name__ == '__main__':
    app.run(debug=True)
