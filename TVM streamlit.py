import streamlit as st
import pandas as pd
from io import BytesIO
import re

def adjust_kenteken(kenteken):

    return re.sub(r'(?<=[A-Z])(?=\d)|(?<=\d)(?=[A-Z])', '-', kenteken)


def process_file(factuur_file):

    factuur = pd.read_excel(factuur_file)

    factuur['Kenteken'] = factuur['Kenteken'].apply(adjust_kenteken)
    
    factuur['Boekjaar'] = factuur['Van'].dt.year
    factuur['Periode'] = factuur['Van'].dt.month

    factuur['Tm'] = factuur['Tot'] - pd.Timedelta(days=1)

    columns = [
        'Dagboek: Code', 'Boekjaar', 'Periode', 'Boekstuknummer', 'Omschrijving: Kopregel',
        'Factuurdatum', 'Vervaldatum', 'Valuta', 'Wisselkoers', 'Betalingsvoorwaarde: Code',
        'Ordernummer', 'Uw ref.', 'Betalingsreferentie', 'Code', 'Naam', 'Grootboekrekening',
        'Omschrijving', 'BTW-code', 'BTW-percentage', 'Bedrag', 'Aantal', 'BTW-bedrag',
        'Opmerkingen', 'Project', 'Van', 'Naar', 'Kostenplaats: Code', 'Kostenplaats: Omschrijving',
        'Kostendrager: Code', 'Kostendrager: Omschrijving'
    ]
    
    import_df = pd.DataFrame(columns=columns)
    import_df['Boekjaar'] = factuur['Boekjaar']
    import_df['Dagboek: Code'] = 60
    import_df['Periode'] = factuur['Periode']
    import_df['Factuurdatum'] = factuur['Van']
    import_df['Uw ref.'] = factuur['Notanummer']
    import_df['Omschrijving: Kopregel'] = factuur['Notanummer'].astype(str) + ' / TVM VERZEKERING'
    import_df['Code'] = 200387
    import_df['Omschrijving'] = factuur['Soort mutatie']
    import_df['Grootboekrekening'] = 7510
    import_df['Bedrag'] = factuur['Nota bedrag']
    import_df['Van'] = factuur['Van']
    import_df['Naar'] = factuur['Tm']
    import_df['Kostenplaats: Code'] = factuur['Kenteken']
    import_df['Kostenplaats: Omschrijving'] = factuur['Kenteken']
    import_df['Factuurdatum'] = import_df['Factuurdatum'].dt.strftime('%d-%m-%Y')
    import_df['Van'] = import_df['Van'].dt.strftime('%d-%m-%Y')
    import_df['Naar'] = import_df['Naar'].dt.strftime('%d-%m-%Y')
    
    new_row = import_df.iloc[0].copy()
    new_row['Bedrag'] = ''
    new_row['Kostenplaats: Code'] = ''
    new_row['Kostenplaats: Omschrijving'] = ''
    import_df = pd.concat([pd.DataFrame([new_row]), import_df], ignore_index=True)

    return import_df


def main():

    st.title('Import TVM factuur')
    
    factuur_file = st.file_uploader('Upload het Excel-factuurbestand', type=['xls', 'xlsx'])
    
    if factuur_file:

        processed_file = process_file(factuur_file)
        st.write('Verwerkte factuur', processed_file.head())
                
        output = BytesIO()
        processed_file.to_excel(output, index=False)

        output.seek(0)

        st.download_button(label='Download verwerkte factuur',
                            data=output,
                            file_name='ImportDaimlerFactuur.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
        st.markdown("<p style='font-size:12px; color:gray;'>Let op, sla het bestand hierna op als Excel 97-2003, zodat je het kan importeren in Exact.</p>", unsafe_allow_html=True)


if __name__ == '__main__':
    main()