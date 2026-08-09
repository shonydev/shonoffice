namespace ShonOffice.Domain.Documentos;

/// <summary>
/// Marca los tipos que pueden aparecer, en cualquier orden, como elemento de
/// primer nivel del cuerpo de un <see cref="DocumentoWord"/>: un
/// <see cref="Parrafo"/> o una <see cref="Tabla"/>. Antes de esto el modelo
/// solo contemplaba párrafos (<c>DocumentoWord.ParrafosConFormato</c>), por
/// lo que un adaptador de lectura que solo recorriera párrafos de nivel
/// superior (<c>cuerpo.Elements&lt;Paragraph&gt;()</c>) ignoraba por
/// completo las tablas: en OOXML una tabla (<c>w:tbl</c>) es un hermano del
/// párrafo dentro de <c>w:body</c>, no un párrafo, así que sus filas nunca
/// aparecían. <see cref="DocumentoWord.ElementosConFormato"/> preserva el
/// orden real del documento entre párrafos y tablas para que la UI pueda
/// reconstruirlo tal como lo muestra Word.
/// </summary>
public interface IElementoDeContenido
{
}
