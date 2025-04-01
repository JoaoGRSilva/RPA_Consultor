import argparse
from contraktor_bot import ContraktorBot

def main():
    parser = argparse.ArgumentParser(description='RPA para processamento de contratos no Contraktor')
    parser.add_argument('--test', action='store_true', help='Modo de teste (processa apenas o primeiro contrato)')
    parser.add_argument('--limit', type=int, help='Limita o número de contratos a processar')
    args = parser.parse_args()
    
    bot = ContraktorBot()
    bot.executar(limite=args.limit, modo_teste=args.test)

if __name__ == "__main__":
    main()