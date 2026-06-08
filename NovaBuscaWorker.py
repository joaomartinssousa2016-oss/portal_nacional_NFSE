import os
from NovaBusca import main


if __name__ == "__main__":
    # Lê variáveis de ambiente passou pela UI
    baixar_xml = os.environ.get("NOVA_BUSCA_BAIXAR_XML", "1") == "1"
    baixar_pdf = os.environ.get("NOVA_BUSCA_BAIXAR_PDF", "0") == "1"
    pdf_tipo = os.environ.get("NOVA_BUSCA_PDF_TIPO", "todos")

    main(baixar_xml=baixar_xml, baixar_pdf=baixar_pdf, pdf_tipo=pdf_tipo)
