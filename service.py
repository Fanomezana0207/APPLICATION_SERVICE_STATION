class Service() :

    def __init__(self, service, vente = 0):
        self.service = service
        self.vente = vente

    def setVente(self, vente) :
        self.vente = vente

    def getService(self) :
        return self.service

    def getVente(self) :
        return self.vente

    def __str__(self):
        return f"{self.service} : {self.vente}"