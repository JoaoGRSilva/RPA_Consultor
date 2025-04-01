class Selectors:
    """Classe para armazenar todos os seletores XPath usados no RPA."""
    
    # Login
    LOGIN_EMAIL = "email"
    LOGIN_PASSWORD = "password"
    LOGIN_BUTTON = "//button[contains(@type, 'submit')]"
    
    # Busca de Contratos
    LOADING_SPINNER = '//div[contains(@class, "ui active transition visible inverted dimmer")]'
    SEARCH_DROPDOWN = '//*[@id="single-spa-application:@contraktor/legacy"]/div/div[2]/div[2]/div/div[2]/div[1]/form/div/div[1]'
    SEARCH_OPTION_NUMERO = "//div[@role='option']//span[contains(@class, 'text') and text()='Número']"
    SEARCH_INPUT = "//input[contains(@type, 'search')]"
    SEARCH_BUTTON = '//*[@id="single-spa-application:@contraktor/legacy"]/div/div[2]/div[2]/div/div[2]/div[1]/form/button'
    
    # Contrato
    CONTRACT_LINK = '//*[@id="number"]/a'
    DOWNLOAD_BUTTON = '//*[@id="single-spa-application:@contraktor/legacy"]/div/div[1]/div[2]/div[1]/div[4]/button[1]'
    RETURN_BUTTON = '//*[@id="single-spa-application:@contraktor/compass"]/header/div/div/nav/div[1]/div[1]/button'