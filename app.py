from flask import Flask, render_template, request, jsonify
import pandas as pd
import os
import json
from tempfile import mkdtemp
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Configurações
UPLOAD_FOLDER = mkdtemp()
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # Limite de 16MB

ALLOWED_EXTENSIONS = {'csv', 'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'Nenhum arquivo selecionado'}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        try:
            # Processar arquivo
            if filename.endswith('.csv'):
                df = pd.read_csv(filepath)
            else:
                df = pd.read_excel(filepath)
            
            # Verificar colunas necessárias
            required_cols = ['MES', 'TIPO_CONS', 'VALOR', 'EMP_CODIGO', 'RUBRICA']
            df_cols = df.columns.str.upper().tolist()
            
            # Mapear colunas
            col_mapping = {}
            for col in required_cols:
                matches = [c for c in df.columns if c.upper() == col]
                if not matches:
                    return jsonify({'error': f'Coluna {col} não encontrada'}), 400
                col_mapping[matches[0]] = col.lower()
            
            # Adicionar FONTE se existir
            fonte_matches = [c for c in df.columns if c.upper() == 'FONTE']
            if fonte_matches:
                col_mapping[fonte_matches[0]] = 'fonte'
                df[fonte_matches[0]] = df[fonte_matches[0]].astype(str)
            else:
                df['FONTE'] = 'N/A'
                col_mapping['FONTE'] = 'fonte'
            
            df = df.rename(columns=col_mapping)
            
            # Aplicar filtro de empresas
            emp_filter = request.form.get('emp_filter')
            if emp_filter:
                try:
                    emp_filter_list = json.loads(emp_filter)
                    if emp_filter_list:
                        df = df[df['emp_codigo'].astype(str).isin(emp_filter_list)]
                except json.JSONDecodeError:
                    pass
            
            # Processar dados
            df['tipo_desc'] = df['tipo_cons'].apply(lambda x: str(x).split(' - ')[1] if ' - ' in str(x) else str(x))
            df['mes'] = df['mes'].astype(str)
            df['month_num'] = df['mes'].str[4:6]
            df = df[df['month_num'].str.isnumeric()]
            df = df[df['month_num'].astype(int).between(1, 12)]
            
            # Resumo por mês
            by_month = df.groupby(['tipo_desc', 'month_num'])['valor'].sum().reset_index()
            pivot_month = by_month.pivot(index='tipo_desc', columns='month_num', values='valor').fillna(0)

            all_months = [f"{m:02d}" for m in range(1, 13)]
            pivot_month = pivot_month.reindex(columns=all_months, fill_value=0)
            
            def convert_to_native_types(obj):
                if isinstance(obj, dict):
                    return {str(k): convert_to_native_types(v) for k, v in obj.items()}
                elif hasattr(obj, 'tolist'):
                    return obj.tolist()
                elif hasattr(obj, 'item'):
                    return obj.item()
                elif isinstance(obj, list):
                    return [convert_to_native_types(i) for i in obj]
                else:
                    return obj
            
            month_summary = convert_to_native_types(pivot_month.to_dict(orient='index'))
            
            # Resumo por fonte
            source_summary = {}
            by_source = df.groupby(['month_num', 'fonte', 'tipo_desc'])['valor'].sum().reset_index()

            for month in all_months:
                source_summary[month] = {}
                month_data = by_source[by_source['month_num'] == month]
                
                all_sources = df['fonte'].unique()
                all_tipos = df['tipo_desc'].unique()
                
                for fonte in all_sources:
                    source_summary[month][str(fonte)] = {str(tipo): 0 for tipo in all_tipos}
                
                for _, row in month_data.iterrows():
                    source_summary[month][str(row['fonte'])][str(row['tipo_desc'])] = float(row['valor'])
                                
            # Resumo por rubricas
            rubricas_summary = {}
            by_rubrica = df.groupby(['month_num', 'tipo_desc', 'rubrica', 'fonte'])['valor'].sum().reset_index()

            for month in all_months:
                rubricas_summary[month] = {}
                month_data = by_rubrica[by_rubrica['month_num'] == month]
                
                for tipo in month_data['tipo_desc'].unique():
                    rubricas_summary[month][str(tipo)] = []
                    tipo_data = month_data[month_data['tipo_desc'] == tipo]
                    
                    for rubrica in tipo_data['rubrica'].unique():
                        rubrica_data = tipo_data[tipo_data['rubrica'] == rubrica]
                        rubrica_entry = {
                            'rubrica': str(rubrica),
                            'fontes': {}
                        }
                        
                        for fonte in all_sources:
                            fonte_value = rubrica_data[rubrica_data['fonte'] == fonte]['valor'].sum()
                            rubrica_entry['fontes'][str(fonte)] = float(fonte_value)
                        
                        rubricas_summary[month][str(tipo)].append(rubrica_entry)
            
            # Preparar resposta
            sources = [str(s) for s in sorted(df['fonte'].unique().tolist())]
            tipos = [str(t) for t in sorted(df['tipo_desc'].unique().tolist())]
            emp_codigos = [str(e) for e in sorted(df['emp_codigo'].unique().tolist())]
            
            sample_data = convert_to_native_types(
                df.head(100)[['mes', 'tipo_cons', 'fonte', 'valor', 'emp_codigo']].to_dict(orient='records')
            )
            
            result = {
                'monthSummary': month_summary,
                'sourceSummary': convert_to_native_types(source_summary),
                'rubricasSummary': convert_to_native_types(rubricas_summary),
                'months': all_months,
                'sources': sources,
                'tipos': tipos,
                'empCodigos': emp_codigos,
                'sampleData': sample_data,
                'totalRows': int(len(df))
            }
            
            return jsonify(result), 200
            
        except Exception as e:
            return jsonify({'error': f'Erro ao processar arquivo: {str(e)}'}), 500
        finally:
            if os.path.exists(filepath):
                os.remove(filepath)
    
    return jsonify({'error': 'Tipo de arquivo não permitido. Use .csv, .xlsx ou .xls'}), 400

if __name__ == '__main__':
    app.run(debug=True)