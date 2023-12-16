# Word_ContarCitasTextuales
Macro en Word para contar las citas textuales en un documento.

El siguiente codigo de Macro en Word 365 sirve para contar las citas textuales que esten en un documento. Esto es util para comparar que las referencias bibliograficas correspondan con las citas textuales directas o indirectas que estan en el documento. El formato utilizado es el estilo APA 7a edicion. Las citas textuales podran estar dentro del parrafo o al final conla estructura:
(Autor, Fecha) o (Autor, Fecha, pag 0).

Ej. (Lozano, 2023) o (Perez, 2021, pag 78)

Es importante notar que en el caso de incluir la pagina de donde se obtuvo la cita esta debe ser con las letras "pag", si se coloca "pg" o "p" solamente, ignorara y no la contara.

El código esta en VBA (Visual Basic for Applications) para Word que cuenta las citas en formato APA en un documento. Puedes seguir estos pasos:

1. Abre tu documento de Word.

2. Presiona Alt + F11 para abrir el Editor de VBA.

3. Inserta un nuevo módulo: Haz clic derecho en "VBAProject (TuDocumento)" -> Insertar -> Módulo.

4. Copia y pega el código en el módulo:

5. Cierra el Editor de VBA.

6. Ahora puedes ejecutar la macro presionando Alt + F8, seleccionando "ContarCitasAPA" y haciendo clic en "Ejecutar".

Este código recorrerá todos los párrafos del documento y contará aquellos que contengan una cita APA. El resultado se mostrará en una caja de diálogo.

