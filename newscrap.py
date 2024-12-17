import scrapy

class HojaViajeraSpider(scrapy.Spider):
    name = 'hojaviajera'
    start_urls = ['https://hojaviajeradigital.javer.com.mx:9260/#/login']

    def parse(self, response):
        # Enviar POST con usuario y contraseña
        return scrapy.FormRequest.from_response(
            response,
            formdata={'correo': 'fprado', 'password': 'pamf900509HFA04'},
            callback=self.after_login
        )

    def after_login(self, response):
        # Verificar si se inició sesión correctamente
        if 'Bienvenido' in response.text:
            self.logger.info("Inicio de sesión exitoso")
            # Acceder a la página de inicio
            yield scrapy.Request('https://hojaviajeradigital.javer.com.mx:9260/#/inicio', callback=self.parse_inicio)

    def parse_inicio(self, response):
        # Acceder a la página de detalle de un folio específico
        folio = '9422'
        yield scrapy.Request(f'https://hojaviajeradigital.javer.com.mx:9260/#/detalle?i={folio}', callback=self.parse_detalle)

    def parse_detalle(self, response):
        # Descargar el archivo zip
        yield scrapy.Request(response.xpath("//mat-icon[text()='download']/parent::button/@ng-reflect-router-link").get(), callback=self.descargar_zip)

    def descargar_zip(self, response):
        # Guardar el archivo zip
        with open('hoja_viajera.zip', 'wb') as f:
            f.write(response.body)
        self.logger.info('Archivo descargado con éxito')