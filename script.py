
import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# Função para carregar o arquivo Excel
def load_excel_file(file_buffer):
    try:
        xls = pd.ExcelFile(file_buffer)
        return xls
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo Excel: {e}")
        return None

# Função para carregar dados de uma aba
@st.cache_data
def load_sheet_data(sheet_name, _xls):
    try:
        df = pd.read_excel(_xls, sheet_name=sheet_name)
        df['Origem'] = sheet_name
        return df
    except Exception as e:
        st.error(f"Erro ao carregar os dados da aba '{sheet_name}': {e}")
        return pd.DataFrame()

# Função para filtros avançados
def advanced_filters(df):
    st.sidebar.subheader("Filtros Avançados")
    
    # Filtro por intervalo de datas
    date_column = st.sidebar.selectbox("Selecione a coluna de data", df.select_dtypes(include=['datetime']).columns)
    if date_column:
        start_date, end_date = st.sidebar.date_input("Selecione o intervalo de datas", [df[date_column].min(), df[date_column].max()])
        df = df[(df[date_column] >= pd.to_datetime(start_date)) & (df[date_column] <= pd.to_datetime(end_date))]
    
    # Filtro por múltiplas colunas
    col_filter = st.sidebar.multiselect("Selecione colunas para filtro", df.columns)
    if col_filter:
        filter_values = {col: st.sidebar.text_input(f"Filtro para {col}", "") for col in col_filter}
        for col, value in filter_values.items():
            if value:
                df = df[df[col].astype(str).str.contains(value, case=False)]
    
    return df

# Função para buscar com filtros
def search_filter_all(df, query, column=None, search_type='Nome'):
    if search_type == 'Nome':
        if query:
            query = query.lower()
            if column and column != "Todas":
                if column in df.columns:
                    mask = df[column].astype(str).str.lower().str.contains(query)
                    return df[mask]
                return df
            else:
                mask = df.apply(lambda row: query in row.astype(str).str.lower().tolist(), axis=1)
                return df[mask]
        return df
    elif search_type == 'Número':
        if query:
            try:
                query = int(query)
                if column and column != "Todas":
                    if column in df.columns:
                        mask = df[column].astype(str).str.contains(str(query))
                        return df[mask]
                    return df
                else:
                    mask = df.apply(lambda row: str(query) in row.astype(str).tolist(), axis=1)
                    return df[mask]
            except ValueError:
                st.error("O valor de busca deve ser um número válido.")
        return df

# Função para buscar em todas as abas
def search_all_sheets(xls, query, search_type, column):
    all_dfs = []
    for sheet_name in xls.sheet_names:
        df = load_sheet_data(sheet_name, xls)
        filtered_df = search_filter_all(df, query, column, search_type)
        if not filtered_df.empty:
            all_dfs.append(filtered_df)
    if all_dfs:
        result_df = pd.concat(all_dfs, ignore_index=True)
        return result_df
    else:
        st.write("Nenhum resultado encontrado.")
        return pd.DataFrame()

# Função para plotar gráfico de pizza com Plotly
def plot_pizza_chart(df, column):
    if column in df.columns:
        fig = px.pie(df, names=column, title=f'Distribuição da coluna {column}', color_discrete_sequence=px.colors.sequential.Blues)
        st.plotly_chart(fig)
    else:
        st.error(f"Coluna '{column}' não encontrada.")

# Função para plotar gráfico de linha com Plotly
def plot_line_chart(df, x_column, y_column):
    if x_column in df.columns and y_column in df.columns:
        fig = px.line(df, x=x_column, y=y_column, title=f'Gráfico de Linha de {y_column} ao longo de {x_column}', color_discrete_sequence=px.colors.sequential.Blues)
        st.plotly_chart(fig)
    else:
        st.error(f"Colunas '{x_column}' ou '{y_column}' não encontradas.")

# Função para exibir dados alterados
def show_edited_data():
    if 'edited_data' in st.session_state and not st.session_state.edited_data.empty:
        st.write("### Dados Alterados")
        st.dataframe(st.session_state.edited_data)
    else:
        st.write("Nenhuma alteração registrada.")

# Função para adicionar uma nova aba
def add_tab_form(xls):
    with st.sidebar.expander("Adicionar Aba", expanded=False):
        new_tab_name = st.text_input("Nome da nova aba", "")
        if st.button("Adicionar Aba", key="add_tab"):
            if new_tab_name and new_tab_name not in xls.sheet_names:
                try:
                    with pd.ExcelWriter("dados_atualizados.xlsx", engine='xlsxwriter') as writer:
                        for sheet_name in xls.sheet_names:
                            df = pd.read_excel(xls, sheet_name=sheet_name)
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                        pd.DataFrame().to_excel(writer, sheet_name=new_tab_name, index=False)
                    st.session_state.sheet_names.append(new_tab_name)
                    st.success("Aba adicionada com sucesso!")
                    st.experimental_rerun()
                except Exception as e:
                    st.error(f"Erro ao adicionar a aba: {e}")
            else:
                st.error("Nome da aba inválido ou já existente.")

# Função para remover uma aba existente
def remove_tab_form(xls):
    with st.sidebar.expander("Remover Aba", expanded=False):
        tab_to_remove = st.selectbox("Selecione a aba para remover", st.session_state.sheet_names, key="remove_tab_select")
        if st.button("Remover Aba", key="remove_tab"):
            if tab_to_remove:
                try:
                    with pd.ExcelWriter("dados_atualizados.xlsx", engine='xlsxwriter') as writer:
                        for sheet_name in xls.sheet_names:
                            if sheet_name != tab_to_remove:
                                df = pd.read_excel(xls, sheet_name=sheet_name)
                                df.to_excel(writer, sheet_name=sheet_name, index=False)
                    st.session_state.sheet_names.remove(tab_to_remove)
                    st.success("Aba removida com sucesso!")
                    st.experimental_rerun()
                except Exception as e:
                    st.error(f"Erro ao remover a aba: {e}")
            else:
                st.error("Selecione uma aba válida para remover.")

# Função para renomear uma aba existente
def rename_tab_form(xls):
    with st.sidebar.expander("Renomear Aba", expanded=False):
        tab_to_rename = st.selectbox("Selecione a aba para renomear", st.session_state.sheet_names, key="rename_tab_select")
        new_tab_name = st.text_input("Novo nome da aba", "", key="rename_tab_input")
        if st.button("Renomear Aba", key="rename_tab"):
            if tab_to_rename and new_tab_name and tab_to_rename in xls.sheet_names and new_tab_name not in xls.sheet_names:
                try:
                    with pd.ExcelWriter("dados_atualizados.xlsx", engine='xlsxwriter') as writer:
                        for sheet_name in xls.sheet_names:
                            df = pd.read_excel(xls, sheet_name=sheet_name)
                            if sheet_name == tab_to_rename:
                                df.to_excel(writer, sheet_name=new_tab_name, index=False)
                            else:
                                df.to_excel(writer, sheet_name=sheet_name, index=False)
                    st.session_state.sheet_names = [new_tab_name if name == tab_to_rename else name for name in st.session_state.sheet_names]
                    st.success("Aba renomeada com sucesso!")
                    st.experimental_rerun()
                except Exception as e:
                    st.error(f"Erro ao renomear a aba: {e}")
            else:
                st.error("Nome da aba inválido ou já existente.")

# Função para editar dados
def edit_data(df):
    if not df.empty:
        with st.sidebar.expander("Editar Dados", expanded=False):
            st.write("### Editar Dados Específicos")
            st.write("#### Editar Linha Existente")
            row_index = st.number_input("Número da linha para editar (0-indexado)", min_value=0, max_value=len(df)-1, step=1, key="edit_row_index")
            if not df.empty:
                col_name = st.selectbox("Coluna para editar", df.columns, key="edit_column_select")
                new_value = st.text_input("Novo valor", value=df.at[row_index, col_name] if not pd.isna(df.at[row_index, col_name]) else "", key="edit_value_input")
                
                if st.button("Salvar Alterações", key="save_edits"):
                    try:
                        df.at[row_index, col_name] = new_value
                        st.session_state.edited_data = df
                        st.success("Alterações salvas com sucesso!")
                    except Exception as e:
                        st.error(f"Erro ao salvar alterações: {e}")

# Função para salvar alterações
def save_changes(df):
    if 'edited_data' in st.session_state and not st.session_state.edited_data.empty:
        try:
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                for sheet_name in st.session_state.sheet_names:
                    if sheet_name in df['Origem'].unique():
                        sheet_df = df[df['Origem'] == sheet_name]
                        sheet_df.drop(columns='Origem', inplace=True)
                        sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
            buffer.seek(0)
            return buffer
        except Exception as e:
            st.error(f"Erro ao salvar alterações: {e}")
            return None
    else:
        st.error("Não há dados para salvar.")
        return None

# Função principal
def main():
    st.title("Dashboard com Manipulação de Dados")
    
    file_buffer = st.file_uploader("Carregar arquivo Excel", type=["xlsx"])
    if file_buffer:
        xls = load_excel_file(file_buffer)
        if xls:
            st.session_state.sheet_names = xls.sheet_names
            
            # Barra lateral para selecionar a aba
            selected_sheet = st.sidebar.selectbox("Selecione a aba", xls.sheet_names)
            df = load_sheet_data(selected_sheet, xls)
            
            st.write("### Dados da Aba Selecionada")
            st.dataframe(df)
            
            # Adicionar filtros avançados
            if st.sidebar.checkbox("Aplicar Filtros Avançados"):
                df = advanced_filters(df)
            
            # Busca e filtros
            st.sidebar.write("### Busca e Filtros")
            query = st.sidebar.text_input("Buscar")
            search_type = st.sidebar.radio("Tipo de busca", ("Nome", "Número"))
            column = st.sidebar.selectbox("Selecionar coluna para busca", options=["Todas"] + df.columns.tolist())
            search_result = search_filter_all(df, query, column if column != "Todas" else None, search_type)
            st.dataframe(search_result)
            
            # Adicionar visualizações
            if st.sidebar.checkbox("Mostrar Gráficos"):
                plot_pizza_chart(df, st.sidebar.selectbox("Coluna para gráfico de pizza", df.columns, key="pizza_column_select"))
                plot_line_chart(df, st.sidebar.selectbox("Coluna X para gráfico de linha", df.columns, key="line_x_column_select"), st.sidebar.selectbox("Coluna Y para gráfico de linha", df.columns, key="line_y_column_select"))
            
            show_edited_data()
            
            # Funções de manipulação de abas
            add_tab_form(xls)
            remove_tab_form(xls)
            rename_tab_form(xls)
            edit_data(df)
            
            # Salvar alterações
            if st.button("Salvar Alterações", key="save_changes"):
                save_result = save_changes(df)
                if save_result:
                    st.download_button("Baixar arquivo atualizado", save_result, "dados_atualizados.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            # Busca em todas as abas
            st.sidebar.write("### Busca em Todas as Abas")
            query_all = st.sidebar.text_input("Buscar em todas as abas")
            search_type_all = st.sidebar.radio("Tipo de busca em todas as abas", ("Nome", "Número"))
            column_all = st.sidebar.selectbox("Selecionar coluna para busca em todas as abas", options=["Todas"] + df.columns.tolist(), key="search_column_all")
            search_result_all = search_all_sheets(xls, query_all, search_type_all, column_all if column_all != "Todas" else None)
            if not search_result_all.empty:
                st.write("### Resultados da Busca em Todas as Abas")
                st.dataframe(search_result_all)

if __name__ == "__main__":
    main()

