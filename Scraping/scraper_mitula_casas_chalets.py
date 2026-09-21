"""Segmento Mitula: casas y chalets en venta en General Pueyrredon.

Reutiliza el scraper original sin modificarlo y escribe sobre los mismos XLSX
cuando se indican las mismas rutas de ``--salida`` y ``--parcial``.
"""

from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

import scraper_mitula as scraper


URL_FILTRADA = (
    "https://casas.mitula.com.ar/find?geoId=R2621716"
    "&operationType=sell&propertyType=house%2Cvilla"
)


def pagina_url_filtrada(pagina: int) -> str:
    """Conserva los filtros de la busqueda y agrega ``page`` cuando corresponde."""
    partes = urlsplit(URL_FILTRADA)
    consulta = dict(parse_qsl(partes.query, keep_blank_values=True))
    if pagina == 1:
        consulta.pop("page", None)
    else:
        consulta["page"] = str(pagina)
    return urlunsplit(
        (partes.scheme, partes.netloc, partes.path, urlencode(consulta), partes.fragment)
    )


def main() -> None:
    scraper.pagina_url = pagina_url_filtrada
    scraper.main()


if __name__ == "__main__":
    main()
