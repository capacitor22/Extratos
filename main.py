import logging
from processaNovos import fProcNovos

def main(tipoProcessamento='Normal'):

    logger = logging.getLogger('root')
    logger.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(levelname)8s:%(asctime)s %(filename)s %(threadName)12s linha:%(lineno)-4d %(message)s')

    #Console Handler
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    ch.setFormatter(formatter)

    #File Handler
    fh = logging.FileHandler('log.log')
    fh.setLevel(logging.INFO)
    fh.setFormatter(formatter)

    logger.addHandler(ch)
    logger.addHandler(fh)

    if (tipoProcessamento == 'Normal'):
        logger.info('tipoProcessamento do if = %s . Processando NOVOS ARQUIVOS.', tipoProcessamento)
        fProcNovos()                
    else:
        logger.info('tipoProcessamento do else = %s', tipoProcessamento)
    
if __name__ == '__main__':
    main()
    # main('Rebuild')