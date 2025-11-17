
class PerfilOcupacionalNoEncontrado(Exception):
    """Excepción lanzada cuando no se encuentra un perfil ocupacional"""
    def __init__(self, nombre_programa):
        self.nombre_programa = nombre_programa
        super().__init__(f"No se encontró perfil para: {nombre_programa}")