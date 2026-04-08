import os
from app.controller.controller import Controller 
import customtkinter as ctk

# Configurações de aparência
ctk.set_appearance_mode("System") 
ctk.set_default_color_theme("blue")

# Inciializa a interface gráfica, define o título, dimensões e contrói todos os widgets (botões, entradas e labels)
class View(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Automação Entradas de Portaria (ATA/PTP)")
        self.geometry("600x450")
        self.att_icon()

        # --- Layout da Interface ---
        self.primeiroContainer = ctk.CTkFrame(self, fg_color="transparent")
        self.primeiroContainer.pack(pady=40)
        self.titulo = ctk.CTkLabel(self.primeiroContainer, text="Automação Entradas de Portaria", font=("Arial", 18, "bold"))
        self.titulo.pack()

        # Container para o diretório PTP
        self.segundoContainer = ctk.CTkFrame(self, fg_color="transparent")
        self.segundoContainer.pack(padx=20, pady=5)
        self.diretorioPTP = ctk.CTkLabel(self.segundoContainer, text="Diretório do PTP", font=("Arial", 15, "bold"))
        self.diretorioPTP.pack(anchor = "w")
        self.entryPTP = ctk.CTkEntry(self.segundoContainer, width=350, placeholder_text="C:/DiretórioPTP")
        self.entryPTP.pack()

        # Container para o diretório ATA
        self.terceiroContainer = ctk.CTkFrame(self, fg_color="transparent")
        self.terceiroContainer.pack(padx=20, pady=5)
        self.diretorioATA = ctk.CTkLabel(self.terceiroContainer, text="Diretório da ATA", font=("Arial", 15, "bold"))
        self.diretorioATA.pack(anchor = "w")
        self.entryATA = ctk.CTkEntry(self.terceiroContainer, width=350, placeholder_text="C:/DiretórioATA")
        self.entryATA.pack()

        # Botão de Execução
        self.quartoContainer = ctk.CTkFrame(self, fg_color="transparent")
        self.quartoContainer.pack(pady=20)
        self.btn_executar = ctk.CTkButton(self.quartoContainer, text="Executar", command=self.executarAut, fg_color="green", hover_color="#006400")
        self.btn_executar.pack(pady=10)

        # Label de Log para feedback ao usuário
        self.quintoContainer = ctk.CTkFrame(self, fg_color="transparent")
        self.quintoContainer.pack(pady=10)
        self.logText = ctk.CTkLabel(self.quintoContainer, text="", font=("Arial", 12, "bold"))
        self.logText.pack()

    # Gerencia o ícone da janela principal, alternando entre as versões Light e Dark dependendo do tema do sistema do usuário
    def att_icon(self):
        # 1. Pegamos o caminho absoluto da pasta onde este script está salvo
        diretorio_atual = os.path.dirname(os.path.realpath(__file__))
        
        # 2. Montamos o caminho até a imagem (independente de onde você rode o script)
        caminho_icone_black = os.path.join(diretorio_atual, "images", "../")
        caminho_icone_white = os.path.join(diretorio_atual, "images", "../")

        mode = ctk.get_appearance_mode()

        try:
            if mode == "Dark" and os.path.exists(caminho_icone_white):
                self.iconbitmap(caminho_icone_white)
            else:
                self.iconbitmap(caminho_icone_black)
        except Exception as e:
            print(f"Erro ao carregar ícone: {e}")

    # Função principal disparada pelo botão 'Executar'. Realiza a validação dos campos, busca os arquivos nos diretórios informados e cahma a lógica do Controller para processar as planilhas    
    def executarAut(self):
        dirptp = self.entryPTP.get()
        dirata = self.entryATA.get()

        # Validação simples de preenchimento
        if not dirptp or not dirata:
           return self.logText.configure(text="Erro: Preencha ambos os campos!", text_color="red")
        
        try:
            self.logText.configure("Buscando arquivos...", text_color="blue")
            self.update()

            # Localiza os arquivos .xlsx nos diretórios fornecidos
            arq_ptp = Controller.pesquisar_arquivo(dirptp)
            arq_ata = Controller.pesquisar_arquivo(dirata)

            # Validação de existência dos arquivos
            if not arq_ptp:
                self.logText.configure(text="Erro: Nenhum arquivo .xlsx no diretório PTP!", text_color="red")
                return
            
            if not arq_ata:
                self.logText.configure(text="Erro: Nenhum arquivo .xlsx no diretório ATA!", text_color="red")
                return
            
            # Atualiza o status visual para o usuário
            self.logText.configure(text=f"Processando: {arq_ptp.name}... \n Processando: {arq_ata.name}...", text_color="orange")
            self.update()

            # Constrói os caminhos absolutos
            ptp_origem = Controller.caminho_origem(dirptp ,arq_ptp)
            ata_destino = Controller.caminho_destino(dirata, arq_ata)

            # Aciona o processo de atualização: Cruza os dados do PTP para preencher as datas na ATA
            Controller.atualizar_ata_com_ptp(ata_destino, ptp_origem)

            # Finalização com sucesso
            self.logText.configure(text="✔ Processo concluído com sucesso!", text_color="green")

        except Exception as e:
            # Caso ocorra algum erro de permissão (arquivo aberto) ou lógica
            self.logText.configure(text=f"Erro crítico: {str(e)}", text_color="red")
            print(f"Detalhes do erro: {e}")

    # Instancia a classe View e inicia o loop principal da interface gráfica.
    def iniciar():
        app = View()
        app.mainloop()
