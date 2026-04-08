from app.models.model import Model


class Controller:

    def pesquisar_arquivo(dict):
       return Model.find_lastet_excel(dict)
    
    def caminho_origem(diretorio_origem ,arquivo_origem):
        return Model.get_file_source(diretorio_origem ,arquivo_origem)
    
    def caminho_destino(diretorio_destino ,arquivo_destino):
        return Model.get_file_destination(diretorio_destino ,arquivo_destino)
    
    def leitura_excel(caminho_origem):
        return Model.excel_data_read(caminho_origem)
    
    def carregar_arquivo(caminho_destino):
        return Model.load_file(caminho_destino)
    
    def converter_df_lista(df_origem):
        return Model.convert_df_list(df_origem)
    
    def inserir_dados(dados_inserir, sheet):
        return Model.insert_data(dados_inserir, sheet)
    
    def salvar_arquivo(book, caminho_destino):
        return Model.save_file(book, caminho_destino)
    
    def atualizar_ata_com_ptp(caminho_destino, caminho_origem):
        return Model.update_ata_with_ptp(caminho_destino, caminho_origem)
    
    