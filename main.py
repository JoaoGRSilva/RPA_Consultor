import argparse
import tkinter as tk
from controllers.contraktor_bot import ContraktorBot  # Importação completa
from views.contraktor_ui import ContraktorBotUI       # Importação completa

def main():
    parser = argparse.ArgumentParser(description='RPA para processamento de contratos no Contraktor')
    parser.add_argument('--test', action='store_true', help='Modo de teste (processa apenas o primeiro contrato)')
    parser.add_argument('--limit', type=int, help='Limita o número de contratos a processar')
    parser.add_argument('--nogui', action='store_true', help='Executa sem interface gráfica (modo CLI)')
    args = parser.parse_args()
    

    if args.nogui:
        bot = ContraktorBot()
        bot.executar(limite=args.limit, modo_teste=args.test)
    else:
        root = tk.Tk()
        app = ContraktorBotUI(root) 
        
        def criar_bot_controller():
            return ContraktorBot(app)
        
        app.criar_bot = criar_bot_controller 
        
        root.mainloop()

if __name__ == "__main__":
    main()