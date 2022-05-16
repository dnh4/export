//-----------------------------------------------------------------------------------<Include>
#pragma once
#include <initFmapApp.h>
#include "export.h"
using namespace std;


//-----------------------------------------------------------------------------------<Fonctions>
/**
  * \fn void cmdExportPoly3d
  * \cmdName EXPORTPOLY3D
  * \desc Exporte les propri�tes des polylignes 3D dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTPOLY3D.gif
  * \bullet Etape 1 : S�lectionner les polylignes 3D.
  * \bullet Etape 2 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \section Lin�aires
  * \end
  */
void cmdExportPoly3d();

/**
  * \fn void cmdExportPoly2d
  * \cmdName EXPORTPOLY2D
  * \desc Exporte les propri�tes des polylignes 2D dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTPOLY2D.gif
  * \bullet Etape 1 : S�lectionner les polylignes 2D.
  * \bullet Etape 2 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \section Lin�aires
  * \end
  */
void cmdExportPoly2d();

/**
  * \fn void cmdExportPoint
  * \cmdName EXPORTPOINT
  * \desc Exporte les propri�tes des points dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTPOINT.gif
  * \bullet Etape 1 : S�lectionner les points.
  * \bullet Etape 2 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \section Ponctuels
  * \end
  */
void cmdExportPoint();

/**
  * \fn void cmdExportEntity
  * \cmdName EXPORTENTITY
  * \desc Exporte les propri�tes des entit�s dans un fichier .xlsx.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTENTITY.gif
  * \bullet Etape 1 : S�lectionner les entit�s.
  * \bullet Etape 2 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \free Remarque : l'export .shp et .txt ne sont pas disponibles avec cette commande.
  * \section G�n�ral
  * \end
  */
void cmdExportEntity();

/**
  * \fn void cmdExportBlock
  * \cmdName EXPORTBLOCK
  * \desc Exporte les propri�tes des blocs dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTBLOCK.gif
  * \bullet Etape 1 : S�lectionner les blocs.
  * \bullet Etape 2 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \section Ponctuels
  * \end
  */
void cmdExportBlock();

/**
  * \fn void cmdExportText
  * \cmdName EXPORTTEXT
  * \desc Exporte les propri�tes des textes dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTTEXT.gif
  * \bullet Etape 1 : S�lectionner les textes.
  * \bullet Etape 2 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \section Ponctuels
  * \end
  */
void cmdExportText();

/**
  * \fn void cmdNewExportMText
  * \cmdName EXPORTMTEXT
  * \desc Exporte les informations des textes multiples dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTMTEXT.gif
  * \bullet Etape 1 : S�lectionner les textes multiples.
  * \bullet Etape 2 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \section Ponctuels
  * \end
  */
void cmdExportMText();

/**
  * \fn void cmdExportPoly3DToXml
  * \cmdName POLY3DTOXML
  * \desc Exporte toutes les polylignes 3D du dessin en xml.
  * \bullet Etape 1 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \section XML
  * \end
  */
void cmdExportPoly3DToXml();


/**
  * \fn void cmdFaceToXML
  * \cmdName -FACETOXML
  * \desc Exporter toutes les faces 3D du dessin en xml.
  * \bullet Etape 1 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \section XML
  * \end
  */
void cmdMoinsFaceToXml();


/**
  * \fn void cmdFaceToXML
  * \cmdName FACETOXML
  * \desc Exporte une s�lection de faces 3D en xml.
  * \brief Commande pour exporter une selection de face en xml et demande le fichier d'export
  * \bullet Etape 1 : S�lectionner les faces 3D.
  * \bullet Etape 2 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \section XML
  * \end
  */
void cmdFaceToXml();

/**
  * \fn void cmdExportPoly
  * \cmdName EXPORTPOLY
  * \desc Exporte les propri�t�s des polylines 2D et des polylines 3D dans des fichiers .txt .csv .xlsx ou .shp.
  * \bullet Etape 1 : S�lectionner les polylignes.
  * \bullet Etape 2 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \section Lin�aires
  * \end
  */
void cmdExportPoly();


/**
  * \fn void cmdExportClosedPoly2d
  * \cmdName EXPORTCLOSEDPOLY2D
  * \desc Exporte les propri�t�s des polylignes 2D ferm�es dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTCLOSEDPOLY2D.gif
  * \bullet Etape 1 : S�lectionner les polylignes 2D ferm�es.
  * \bullet Etape 2 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \free Remarque : S'il s'agit d'un export .shp, la couche d'export est polygonale.
  * \section Polygones
  * \end
  */
void cmdExportClosedPoly2d();


/**
  * \fn void cmdExportClosedPoly3d
  * \cmdName EXPORTCLOSEDPOLY3D
  * \desc Exporte les propri�t�s des polylignes 3D ferm�es dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTCLOSEDPOLY3D.gif
  * \bullet Etape 1 : S�lectionner les polylignes 3D ferm�es.
  * \bullet Etape 2 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \free Remarque : S'il s'agit d'un export .shp, la couche d'export est polygonale.
  * \section Polygones
  * \end
  */
void cmdExportClosedPoly3d();


/**
  * \fn void cmdExportClosedPoly
  * \cmdName EXPORTCLOSEDPOLY
  * \desc Exporte les propri�t�s des polylignes 2D et des polylignes 3D ferm�es dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTCLOSEDPOLY.gif
  * \bullet Etape 1 : S�lectionner les polylignes ferm�es.
  * \bullet Etape 2 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \free Remarque : S'il s'agit d'un export .shp, la couche d'export est polygonale.
  * \section Polygones
  * \end
  */
void cmdExportClosedPoly();

/**
  * \fn void cmdExportLine
  * \cmdName EXPORTLINE
  * \desc Exporte les propri�t�s des lignes dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTLINE.gif
  * \bullet Etape 1 : S�lectionner les lignes.
  * \bullet Etape 2 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \section Lin�aires
  * \end
  */
void cmdExportLine();

/**
  * \fn void cmdExportCurve
  * \cmdName EXPORTCURVE
  * \desc Exporte les propri�t�s des polylignes et des lignes dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTCURVE.gif
  * \bullet Etape 1 : S�lectionner les polylignes et les lignes.
  * \bullet Etape 2 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \section Lin�aires
  * \end
  */
void cmdExportCurve();

/**
  * \fn void cmdExportCircle
  * \cmdName EXPORTCIRCLE
  * \desc Exporte les propri�t�s des cercles dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTCIRCLE.gif
  * \bullet Etape 1 : S�lectionner les cercles.
  * \bullet Etape 2 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \section Ponctuels
  * \end
  */
void cmdExportCircle();

/**
  * \fn void cmdExportBlockPoly
  * \cmdName EXPORTBLOCKPOLY
  * \desc Exporte les propri�t�s des blocs situ�s � l'int�rieur d'une polyligne dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTBLOCKPOLY.gif
  * \bullet Etape 1 : S�lectionner les blocs.
  * \bullet Etape 2 : S�lectionner les polylignes.
  * \bullet Etape 3 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \free Remarque : S'il s'agit d'un export .shp, la couche d'export est polygonale.
  * \section Polygones
  * \end
  */
void cmdExportBlockPoly();

/**
  * \fn void cmdExportBlockPoly2D
  * \cmdName EXPORTBLOCKPOLY2D
  * \desc Exporte les propri�t�s des blocs situ�s � l'int�rieur d'une polyligne 2D dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTBLOCKPOLY2D.gif
  * \bullet Etape 1 : S�lectionner les blocs.
  * \bullet Etape 2 : S�lectionner les polylignes 2D.
  * \bullet Etape 3 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \free Remarque : S'il s'agit d'un export .shp, la couche d'export est polygonale.
  * \section Polygones
  * \end
  */
void cmdExportBlockPoly2D();

/**
  * \fn void cmdExportBlockPoly3D
  * \cmdName EXPORTBLOCKPOLY3D
  * \desc Exporte les propri�t�s des blocs situ�s � l'int�rieur d'une polyligne 3D dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTBLOCKPOLY3D.gif
  * \bullet Etape 1 : S�lectionner les blocs.
  * \bullet Etape 2 : S�lectionner les polylignes 3D.
  * \bullet Etape 3 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \free Remarque : S'il s'agit d'un export .shp, la couche d'export est polygonale.
  * \section Polygones
  * \end
  */
void cmdExportBlockPoly3D();

/**
  * \fn void cmdExportTextPoly2D
  * \cmdName EXPORTTEXTPOLY2D
  * \desc Exporte les propri�t�s des textes situ�s � l'int�rieur d'une polyligne 2D dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTTEXTPOLY2D.gif
  * \bullet Etape 1 : S�lectionner les textes.
  * \bullet Etape 2 : S�lectionner les polylignes 2D.
  * \bullet Etape 3 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \free Remarque : S'il s'agit d'un export .shp, la couche d'export est polygonale.
  * \section Polygones
  * \end
  */
void cmdExportTextPoly2D();

/**
  * \fn void cmdExportTextPoly3D
  * \cmdName EXPORTTEXTPOLY3D
  * \desc Exporte les propri�t�s des textes situ�s � l'int�rieur d'une polyligne 3D dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTTEXTPOLY3D.gif
  * \bullet Etape 1 : S�lectionner les textes.
  * \bullet Etape 2 : S�lectionner les polylignes 3D.
  * \bullet Etape 3 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \free Remarque : S'il s'agit d'un export .shp, la couche d'export est polygonale.
  * \section Polygones
  * \end
  */
void cmdExportTextPoly3D();

/**
  * \fn void cmdExportTextPoly
  * \cmdName EXPORTTEXTPOLY
  * \desc Exporte les propri�t�s des textes situ�s � l'int�rieur d'une polyligne dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTTEXTPOLY.gif
  * \bullet Etape 1 : S�lectionner les textes.
  * \bullet Etape 2 : S�lectionner les polylignes.
  * \bullet Etape 3 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \free Remarque : S'il s'agit d'un export .shp, la couche d'export est polygonale.
  * \section Polygones
  * \end
  */
void cmdExportTextPoly();

/**
  * \fn void cmdExportPonctuel
  * \cmdName EXPORTPONCTUEL
  * \desc Exporte les propri�t�s des ponctuels (points, blocs, cercles, textes et textes multiples) dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTPONCTUEL.gif
  * \bullet Etape 1 : S�lectionner les ponctuels.
  * \bullet Etape 2 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \section Ponctuels
  * \end
  */
void cmdExportPonctuel();

/**
  * \fn void cmdExportHatch
  * \cmdName EXPORTHATCH
  * \desc Exporte les propri�t�s des hachures dans des fichiers .txt .csv .xlsx ou .shp.
  * \imgPath C:\Futurmap\Outils\GStarCAD2020\HTM\images\EXPORTHATCH.gif
  * \bullet Etape 1 : S�lectionner les hachures.
  * \bullet Etape 2 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \section Polygones
  * \end
  */
void cmdExportHatch();

/**
  * \fn void cmdExportAtt
  * \cmdName EXPORTATT
  * \desc Exporte les attributs d'un bloc dans des fichiers .txt ou .csv.
  * \bullet Etape 1 : S�lectionner les blocs dont on veut exporter les attributs.
  * \bullet Etape 2 : Entrer le r�pertoire, le nom et l'extension du fichier pour l'enregistrement.
  * \section Attributs
  * \end
  */
void cmdExportAtt();

/**
  * \fn void cmdExportRGBColor
  * \cmdName EXPORTRGBCOLOR
  * \desc Exporte les couleurs rgb dans un fichier excel ==> indice | rouge | vert | bleu
  * \bullet Etape 1 : Lancer la commande EXPORTRGBCOLOR.
  * \bullet Etape 2 : Entrer le r�pertoire, le nom du fichier pour l'enregistrement.
  * \section Couleurs
  * \end
  */
void cmdExportRGBColor();
