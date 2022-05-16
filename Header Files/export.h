/**
 *  @file template.h
 *  @author Romain LEMETTAIS
 *  @brief Fonctions correspondants au coeur des commandes
 */


//-----------------------------------------------------------------------------------<Include>
#pragma once
#include <adslib.h>
#include <tchar.h>
#include <vector>
#include <initFmapApp.h>
#include "helperFunctions.h"
#include "C:/Futurmap/Dev-Outils/GStarCAD/GRX20/fmaplib/Header Files/acString.h"
#include "C:\Futurmap\Dev-Outils\GStarCAD\GRX20\shapelib-1.5.0\shapefil.h"

#include "blockEntity.h"
#include "pointEntity.h"
#include "poly3dEntity.h"
#include "layer.h"
#include "file.h"
#include "userInput.h"
#include "entityLib.h"
#include "C:\Futurmap\Dev-Outils\GStarCAD\GRX20\shapelib-1.5.0\shapefil.h"

#define PRJ_PATH _T("C:\Futurmap\Outils\GStarCAD2020\TXT\Luxembourg_1930_Gauss.prj")

/*==> Utils function shape <== */

/*Constantes permettant de gérer autant d'attributs d'intérêt que l'on souhaite dans
std::vector<int> attributesId. Ex : attributesId[BLOCK] renvoie le numéro de la
colonne contenant l'attribut BLOCK dans le shapefile.
*/
const int BLOCK = 0;
const int LAYER = 1;
const int COLOR = 2;
const int LINETYPE = 3;
const int LINESCALE = 4;
const int LINEWEIGHT = 5;
const int CONSTANTWIDTH = 6;
const int ANGLE = 7;
const int SCALE = 8;
const int XSCALE = 9;
const int YSCALE = 10;
const int ZSCALE = 11;
const int HATCH = 12;

// Si rajout d'un attribut d'intérêt ci-dessus, incrémenter NB_ATT
const int NB_ATT = 13;

// Tableau des possibilités de langage
static char* attOptions[NB_ATT][5] =
{
    {"BLOC", "BLOCK", "", "", ""},
    {"CALQUE", "LAYER", "", "", "" },
    {"COULEUR", "COLOR", "", "", ""},
    {"TYPELIGNE", "LINETYPE", "LTYPE", "", ""},
    {"LINESCALE", "LTSCALE", "ECHELLETYPELIGNE", "ECHELLELIGNE", "ECHELLETL"},
    {"LINEWEIGHT", "LWEIGHT", "EPAISSEURLIGNE", "EPLIGNE", ""},
    {"CONSTANTWIDTH", "WIDTH", "CWIDTH", "LARGEURCONSTANTE", "LARGEUR"},
    {"ANGLE", "ORIENTATION", "", "", ""},
    {"ECHELLE", "SCALE", "", "", ""},
    {"XSCALE", "ECHELLEX", "", "", ""},
    {"YSCALE", "ECHELLEY", "", "", ""},
    {"ZSCALE", "ECHELLEZ", "", "", ""},
    {"PATTERN", "HATCH", "MOTIF", "", ""}
};


/**
    \brief Pour chaque attribut, renvoie l'indice de la colonne du shapefile dans laquelle il a été détecté et -1 s'il n'est pas détecté
    \brief Chaque attribut est préalablement défini dans le header avec un set de versions du nom dans attOptions
    \param hDBF Accès au fichier shapefile
    \param pnEntities Nombre de champs du fichier
    \param  attributesId Tableau où chaque indice correspond à un attribut et dont la valeur de la case est le numéro de la colonne correspondante dans le shapefile
    Par exemple, si BLOC = 0 et attributesId[0] = 2, l'attribut bloc a été détecté dans la colonne 2 du shapefile
    \param std::string configPath Chemin vers le fichier de config (par défaut vide)
    \param defaultValues Vecteur contenant les valeurs par défaut des différents paramètres (même indexation que attributesId)
    */
void getFieldsInfo( const DBFHandle& hDBF,
    int( &attributesId )[NB_ATT],
    const AcString& configPath = AcString::kEmpty,
    std::vector<string>& defaultValues = std::vector<string>( 0 ) );

/**
    \brief Pour chaque attribut, renvoie l'indice de la colonne du shapefile dans laquelle il a été détecté et -1 s'il n'est pas détecté
    \brief Chaque attribut est préalablement défini dans le header avec un set de versions du nom dans attOptions
    \param hDBF Accès au fichier shapefile
    \param pnEntities Nombre de champs du fichier
    \param  attributesId Tableau où chaque indice correspond à un attribut et dont la valeur de la case est le numéro de la colonne correspondante dans le shapefile
    Par exemple, si BLOC = 0 et attributesId[0] = 2, l'attribut bloc a été détecté dans la colonne 2 du shapefile
*/
void getAttFieldIndex( const DBFHandle& hDBF,
    int( &attributesId )[NB_ATT] );

/**
    \brief Met à jour le calque d'une entité en fonction du nom du calque
    \param hDBF Accès au fichier shapefile
    \param pEnt Pointeur vers l'entité
    \param i Indice de l'entité dans la table QGIS
    \param layerFieldIndex Colonne du champ calque dans la table QGIS (-1 si le champ n'existe pas)
    \return ErrorStatus
*/
Acad::ErrorStatus setLayer( const DBFHandle& hDBF,
    AcDbEntity* pEnt,
    const int& i,
    const int& layerFieldIndex );

/**
    \brief Met à jour la couleur d'une entité en fonction d'un indice couleur ou d'un triplet RGB
    \param hDBF Accès au fichier shapefile
    \param pEnt Pointeur vers l'entité
    \param i Indice de l'entité dans la table QGIS
    \param colorFieldIndex Colonne du champ couleur dans la table QGIS (-1 si le champ n'existe pas)
    \return ErrorStatus
*/
Acad::ErrorStatus setColor( const DBFHandle& hDBF,
    AcDbEntity* pEnt,
    const int& i,
    const int& colorFieldIndex );

/**
    \brief Met à jour le type de ligne d'une entité
    \param hDBF Accès au fichier shapefile
    \param pEnt Pointeur vers l'entité
    \param i Indice de l'entité dans la table QGIS
    \param linetypeIndex Colonne du champ type de ligne dans la table QGIS (-1 si le champ n'existe pas)
    \return ErrorStatus
*/
Acad::ErrorStatus setLinetype( const DBFHandle& hDBF,
    AcDbEntity* pEnt,
    const int& i,
    const int& linetypeFieldIndex );

/**
    \brief Met à jour l'échelle de type de ligne d'une entité
    \param hDBF Accès au fichier shapefile
    \param pEnt Pointeur vers l'entité
    \param i Indice de l'entité dans la table QGIS
    \param linescaleFieldIndex Colonne du champ échelle de type de ligne dans la table QGIS (-1 si le champ n'existe pas)
    \return ErrorStatus
*/
Acad::ErrorStatus setLinescale( const DBFHandle& hDBF,
    AcDbEntity* pEnt,
    const int& i,
    const int& linescaleFieldIndex );

/**
    \brief Met à jour l'épaisseur de type de ligne d'une entité
    \param hDBF Accès au fichier shapefile
    \param pEnt Pointeur vers l'entité
    \param i Indice de l'entité dans la table QGIS
    \param lineweightFieldIndex Colonne du champ épaisseur de type de ligne dans la table QGIS (-1 si le champ n'existe pas)
    \return ErrorStatus
*/
Acad::ErrorStatus setLineweight( const DBFHandle& hDBF,
    AcDbEntity* pEnt,
    const int& i,
    const int& lineweightFieldIndex );

/**
    \brief Met à jour la largeur constante d'une polyligne 2D
    \param hDBF Accès au fichier shapefile
    \param poly Pointeur vers la polyligne 2D
    \param i Indice de l'entité dans la table QGIS
    \param constantwidthFieldIndex Colonne du champ largeur constante dans la table QGIS (-1 si le champ n'existe pas)
    \return ErrorStatus
*/
Acad::ErrorStatus setConstantWidth( const DBFHandle& hDBF,
    AcDbPolyline* poly,
    const int& i,
    const int& constantwidthFieldIndex );

/**
    \brief Met à jour les hachures d'une surface 3D
    \param hDBF Accès au fichier shapefile
    \param poly Vecteur de pointeurs vers les polylignes représentant les bordures de la hachure
    \param i Indice de l'entité dans la table QGIS
    \param hatchFieldIndex Colonne du champ hachure dans la table QGIS (-1 si le champ n'existe pas)
    \return ErrorStatus
*/
Acad::ErrorStatus setHatch3d( const DBFHandle& hDBF,
    vector<AcDb3dPolyline*> poly,
    const int& i,
    const int& hatchFieldIndex,
    int( &attributesId )[NB_ATT] );

/**
    \brief Met à jour les hachures d'une surface 2D
    \param hDBF Accès au fichier shapefile
    \param poly Vecteur de pointeurs vers les polylignes représentant les bordures de la hachure
    \param i Indice de l'entité dans la table QGIS
    \param hatchFieldIndex Colonne du champ hachure dans la table QGIS (-1 si le champ n'existe pas)
    \return ErrorStatus
*/
Acad::ErrorStatus setHatch2d( const DBFHandle& hDBF,
    vector<AcDbPolyline*> poly,
    const int& i,
    const int& hatchFieldIndex,
    int( &attributesId )[NB_ATT] );

/**
    \brief Met à jour les attributs d'un bloc en fonction de ceux de la table QGIS
    \param pBlock Pointeur vers la référence de bloc
    \param hDBF Accès au fichier shapefile
    \param i Indice de l'entité dans la table QGIS
    \param poly ???
    \return ErrorStatus
*/
Acad::ErrorStatus setAttributes( AcDbBlockReference* pBlock,
    const DBFHandle& hDBF,
    const int& i,
    AcDbEntity* poly );

Acad::ErrorStatus setProperties( AcDbBlockReference*& pBlock,
    const DBFHandle& hDBF,
    const int& i,
    AcDbEntity* poly );

/**
    \brief Met à jour les caractéristiques communes à toutes les entités importées (calque, couleur, type de ligne, échelle de type de ligne et épaisseur de type de ligne)
    \param hDBF Accès au fichier shapefile
    \param pEnt Pointeur vers la référence de bloc
    \param i Indice de l'entité dans la table QGIS
    \param attributesId Tableau où chaque indice correspond à un attribut et dont la valeur de la case est le numéro de la colonne correspondante dans le shapefile
    \return ErrorStatus
*/
Acad::ErrorStatus setCharacteristics( const DBFHandle& hDBF,
    AcDbEntity* pEnt,
    const int& i,
    int( &attributesId )[NB_ATT] );


Acad::ErrorStatus askBlockName( const SHPHandle& hSHP,
    const DBFHandle& hDBF,
    const int& blockFieldIndex,
    AcString& blockName );


Acad::ErrorStatus insertBlock( const DBFHandle& hDBF,
    SHPObject* iShape,
    const int& i,
    const AcGePoint3d& ptInsert,
    const AcString& blockName,
    int( &attributesId )[NB_ATT],
    AcDbEntity* poly = NULL );


Acad::ErrorStatus importShp( const SHPHandle& hSHP,
    const DBFHandle& hDBF,
    const int& pnShapeType,
    const int& pnEntities,
    int( &attributesId )[NB_ATT],
    AcString blockName,
    bool drawHatch );

Acad::ErrorStatus importPoints( const SHPHandle& hSHP,
    const DBFHandle& hDBF,
    const int& pnEntities,
    int( &attributesId )[NB_ATT],
    const int& blockFieldIndex,
    AcString blockName );


Acad::ErrorStatus importAcDbPolyline3d( const SHPHandle& hSHP,
    const DBFHandle& hDBF,
    const int& pnEntities,
    int( &attributesId )[NB_ATT],
    const int& blockFieldIndex,
    AcString blockName );


Acad::ErrorStatus importAcDbPolyline( const SHPHandle& hSHP,
    const DBFHandle& hDBF,
    const int& pnEntities,
    int( &attributesId )[NB_ATT],
    const int& blockFieldIndex,
    AcString blockName );


Acad::ErrorStatus importPolygons3d( const SHPHandle& hSHP,
    const DBFHandle& hDBF,
    const int& pnEntities,
    int( &attributesId )[NB_ATT],
    const int& blockFieldIndex,
    AcString blockName,
    bool drawHatch );

Acad::ErrorStatus importPolygons2d( const SHPHandle& hSHP,
    const DBFHandle& hDBF,
    const int& pnEntities,
    int( &attributesId )[NB_ATT],
    const int& blockFieldIndex,
    AcString blockName,
    bool drawHatch );

/**
    \brief Met à jour la valeur de l'angle de rotation et des échelles en fonction des valeurs des champs associés
    \param hDBF Accès au fichier shapefile
    \param i Indice de l'entité ouverte
    \param attributesId Tableau où chaque indice correspond à un attribut et dont la valeur de la case est le numéro de la colonne correspondante dans le shapefile
    \param angle Valeur de l'angle à mettre à jour
    \param scaleX Valeur de l'échelle en X à mettre à jour
    \param scaleY Valeur de l'échelle en Y à mettre à jour
    \param scaleZ Valeur de l'échelle en Z à mettre à jour
*/
void getAngleAndScale(
    const DBFHandle& hDBF,
    const int& i,
    int( &attributesId )[NB_ATT],
    double& angle,
    double& scaleX,
    double& scaleY,
    double& scaleZ
);

void convertShpToDwg(
    char* filePathShp,
    AcString dir,
    AcString name,
    AcString ext
);


void convertShpToDwg(
    char* filePathShp,
    AcString dir,
    AcString name,
    AcString ext
);

vector <AcString> readParamFile(
    const AcString&
);

int createFields(
    DBFHandle&,
    vector <AcString>
);

void writeEntHandle(
    AcDbEntity* ent,
    DBFHandle dbfHandle,
    int iShape, AcString = _T( "HANDLE" )
);

void writeStatus(
    AcDbEntity* ent,
    DBFHandle dbfHandle,
    int iShape,
    AcString = _T( "CLOSED" )
);

void writeLength2d( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString = _T( "LENGTH2D" ) );

void writeLength3d( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString = _T( "LENGTH3D" ) );

void writeMaterialName( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString = _T( "MATERIAL" ) );

void writeLayerAttribute( AcDbEntity*, DBFHandle, int, AcString = _T( "LAYER" ) );

void writeColorAttribute( AcDbEntity*, DBFHandle, int, AcString = _T( "COLOR" ) );

void writePolyHandle( AcDbEntity*, DBFHandle, int, AcString = _T( "POLYHANDLE" ) );

void writeLynetypeAttribute( AcDbEntity*, DBFHandle, int, AcString = _T( "LINETYPE" ) );

void writeLineTypeScaleAttribute( AcDbEntity*, DBFHandle, int, AcString = _T( "LTSCALE" ) );

void writeLineWeightAttribute( AcDbEntity*, DBFHandle, int, AcString = _T( "LWEIGHT" ) );

void writeTransparency( AcDbEntity*, DBFHandle, int, AcString = _T( "TRANSPARENCY" ) );

void writeBlockName( AcDbEntity*, DBFHandle, int, AcString = _T( "BLOCK" ) );

void writeOrientation( AcDbEntity*, DBFHandle, int, AcString = _T( "ANGLE" ) );

void writeXPos( AcDbEntity*, DBFHandle, int, AcString = _T( "POSX" ) );
void writeXScale( AcDbEntity*, DBFHandle, int, AcString = _T( "XSCALE" ) );

void writeYPos( AcDbEntity*, DBFHandle, int, AcString = _T( "POSY" ) );
void writeYScale( AcDbEntity*, DBFHandle, int, AcString = _T( "YSCALE" ) );

void writeZPos( AcDbEntity*, DBFHandle, int, AcString = _T( "POSZ" ) );

void writeArea( AcDbEntity*, DBFHandle, int, AcString = _T( "AREA" ) );
void writePattern( AcDbHatch*, DBFHandle, int, AcString = _T( "PATTERN" ) );

void writeThicknessAttribute( AcDbEntity*, DBFHandle, int, AcString = _T( "THICKNESS" ) );
void writeLineStartX( AcDbEntity*, DBFHandle, int, AcString = _T( "STARTX" ) );
void writeLineStartY( AcDbEntity*, DBFHandle, int, AcString = _T( "STARTY" ) );
void writeLineStartZ( AcDbEntity*, DBFHandle, int, AcString = _T( "STARTZ" ) );

void writeLineEndX( AcDbEntity*, DBFHandle, int, AcString = _T( "ENDX" ) );
void writeLineEndY( AcDbEntity*, DBFHandle, int, AcString = _T( "ENDY" ) );
void writeLineEndZ( AcDbEntity*, DBFHandle, int, AcString = _T( "ENDZ" ) );
void writeRotation( AcDbEntity*, DBFHandle, int, AcString = _T( "ROTATION" ) );

void writeZScale( AcDbEntity*, DBFHandle, int, AcString = _T( "ZSCALE" ) );

void writeBlock( AcDbEntity*, DBFHandle, int, vector<AcString> );

void writeBlockPoly( AcDbEntity*, AcDbEntity*, DBFHandle, int, vector<AcString> );

void writeAttribute( AcDbEntity*, DBFHandle, int, vector<AcString> );

void writePattern( AcDbHatch*, DBFHandle, int, AcString );

void writeConsantWidht( AcDbEntity*, DBFHandle, int, AcString = _T( "CWIDTH" ) );

void writeLine( AcDbEntity*, DBFHandle, const int&, const vector<AcString>& );
void writeCurve( AcDbEntity*, DBFHandle, const int&, const vector<AcString>& );
void writePoly2d( AcDbEntity*, DBFHandle, const int&, const vector<AcString>& );
void writeCircle( AcDbEntity*, DBFHandle, const int&, const vector<AcString>& );
void writePoly3d( AcDbEntity*, DBFHandle, const int&, const vector<AcString>& );
void writeClosedPoly2d( AcDbEntity*, DBFHandle, const int&, const vector<AcString>& );
void writeClosedPoly3d( AcDbEntity*, DBFHandle, const int&, const vector<AcString>& );
void writeHatch( AcDbHatch*, DBFHandle, const int&, const vector<AcString>& );
void writeTextStr( AcDbEntity*, DBFHandle, int, AcString = _T( "VALUE" ) );

void writeTexteHeight( AcDbEntity*, DBFHandle, int, AcString = _T( "HEIGHT" ) );

void writeTextAlign( AcDbEntity*, DBFHandle, int, AcString = _T( "ALIGN" ) );
void writeRadius( AcDbEntity*, DBFHandle, int, AcString = _T( "RADIUS" ) );

void writeTextStyle( AcDbEntity*, DBFHandle, int, AcString = _T( "STYLE" ) );

void writePoint( AcDbEntity*, DBFHandle, const int&, const vector<AcString>& );

void drawBlockShp( AcDbBlockReference*, SHPHandle&, DBFHandle&, vector<AcString>& );

vector<AcString> createBlockField( const AcString& acs = "" );

vector<AcString> createBlockPoly2dField( const AcString& acs = "" );

vector<AcString> createBlockPoly3dField( const AcString& acs = "" );

vector<AcString> createLineField( const AcString& acs = "" );

vector<AcString> createPolyField( const AcString& acs = "" );

vector<AcString> createClosedPolyField( const AcString& acs = "" );

vector<AcString> createCircleField( const AcString& acs = "" );

vector<AcString> createPoly3dField( const AcString& acs = "" );

vector<AcString> createClosedPoly3dField( const AcString& acs = "" );

vector<AcString> createPointField( const AcString& acs = "" );

vector<AcString> createCurveField( const AcString& acs = "" );

vector<AcString> createHatchField( const AcString& acs = "" );

void drawLineShp( AcDbLine*, SHPHandle&, DBFHandle&, vector<AcString>& );

void drawLineShp( AcDbPolyline*, SHPHandle&, DBFHandle&, vector<AcString>& );

void drawPonctuel( AcDbEntity*, SHPHandle&, DBFHandle&, vector<AcString>& );

void drawCurve( AcDbEntity*, SHPHandle&, DBFHandle&, vector<AcString>& );

void drawCircleShp( AcDbCircle*, SHPHandle&, DBFHandle&, vector<AcString>& );

void drawClosedPolyShp( AcDbPolyline*, vector<AcGePoint3d>& pts, SHPHandle&, DBFHandle&, vector<AcString>& );

void drawClosedPolyShp( AcDb3dPolyline*, vector<AcGePoint3d>& pts, SHPHandle&, DBFHandle&, vector<AcString>& );

void drawLineShp( AcDb3dPolyline*, SHPHandle&, DBFHandle&, vector<AcString>& );

void drawSurfaceShp( AcDbPolyline*, SHPHandle&, DBFHandle&, vector<AcString>& );

void drawSurfaceShp( AcDb3dPolyline*, SHPHandle&, DBFHandle&, vector<AcString>& );

void drawHatch( AcDbHatch*, SHPHandle&, DBFHandle&, vector<AcString>& );

vector<AcString> createSurfaceField( const AcString& acs = "" );

void drawPointShp( AcDbPoint*, SHPHandle&, DBFHandle&, vector<AcString>& );
void drawHatchShp( AcDbHatch*, SHPHandle&, DBFHandle&, vector<AcString>& );

vector<AcString> createTextField( const AcString& acs = "" );

vector<AcString> createTextPoly( const AcString& acs = "" );

vector<AcString> createTextPoly3d( const AcString& acs = "" );

void drawTextShp( AcDbText* txt,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field );

void drawMTextShp( AcDbMText* txt,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field );

void writeText( AcDbEntity* ent,
    DBFHandle dBFHandle,
    int iShape,
    const vector<AcString>& fieldName );

void writeTextPoly( AcDbEntity* ent,
    AcDbEntity* entPoly,
    DBFHandle dBFHandle,
    int iShape,
    const vector<AcString>& fieldName );

void drawTextPoly( AcDbText* txt,
    AcDbEntity* id,
    vector<AcGePoint3d> pt,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field );

void drawblocSurfacekShp( AcDbBlockReference* block,
    vector<AcGePoint3d> pt,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field );


void drawBlockPoly( AcDbBlockReference* block,
    AcDbEntity* polyEnt,
    vector<AcGePoint3d> pt,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field );

void drawblocSurfacekShp( AcDbText* block,
    vector<AcGePoint3d> pt,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field );

bool isSurfaceIside( AcGePoint3dArray surf1, AcGePoint3dArray surf2 );

/*==> End Utils function shape <== */

struct LinePoly
{
    //Premier point de la ligne
    AcGePoint3d ptStart;
    
    //Second point de la ligne
    AcGePoint3d ptEnd;
    
    //Distance entre les deux points
    double curveDistance;
};


struct FacesInLayer
{
    //Calque des faces
    AcString layer;
    
    //Vecteur d'objectId de face
    vector<AcDbObjectId> facesIds;
};

//Encode excel filepath into utf8 encodage
string latin1_to_utf8( const string& latin1 );

//-----------------------------------------------------------------------------------<Fonctions>


/**
  * \brief Fonction pour prendre les propriétes communs des entités
  * \param pl entité principal
  * \param entityLayerType Type de calque
  * \param entityLineType Type de ligne
  * \param entityMaterialName Nom du materiel
  * \param entityColor Couleur
  * \param entityLineWeigth Taille de ligne
  * \param entityTransp Transparence
  * \param entityLineTypeScale Echelle de type de ligne
  * \param enityPlotStyleName Nom du style de tracé
  * \return Valeur des propriétes en communs des entités
  */
void getEntityParams(
    const AcDbEntity* ent,
    AcString& entityLayerType,
    AcString& entityLineType,
    AcString& entityMaterialName,
    AcString& entityColor,
    AcString& entityHandle,
    int& entityLineWeigth,
    AcString& entityLineWeigthStr,
    AcCmTransparency& entityTransp,
    AcString& entityTranspStr,
    double& entityLineTypeScale,
    AcString& entityPlotStyleName );

/**
  * \brief Fonction pour exporter les propriétes des polylignes 2D
  * \param pl2d L'entité polyligne 2d
  * \param pData liste des informations de chaque polyligne 2D en AcStringArray
  * \pData ("Handle|Calque|Longueur|Type de ligne|Linetype Scale|Matériel|Taille de ligne|Couleur|Transparence|Fermée|Thickness|Nom du style de tracé")
  * \return Valeur à exporter des paramètres des polylignes 2d
  */
void exportPoly2d(
    AcDbPolyline*& pl2d,
    AcStringArray& pData
);

void exportPoly(
    AcDbEntity*& ent,
    AcStringArray& pData
);

/**
  * \brief Fonction pour exporter les propriétes des polylignes 3D
  * \param pl3d L'entité polyligne 3d
  * \param pData liste des informations de chaque polyligne 3D en AcStringArray
  * \pData ("Handle|Calque|Longueur|Longueur 2d|Type de ligne|Linetype Scale|Matériel|Taille de ligne|Couleur|Transparence|Fermée|Nom du style de tracé")
  * \return Valeur à exporter des paramètres des polylignes 3d
  */
void exportPoly3d(
    AcDb3dPolyline*& pl3d,
    AcStringArray& pData
);

/**
  * \brief Fonction pour exporter les propriétes des polylignes fermées
  * \param ent L'entité polyligne fermée
  * \param pData liste des informations de chaque polyligne en AcStringArray
  * \pData ("Handle|Calque|Longueur|Longueur 2d|Type de ligne|Linetype Scale|Matériel|Taille de ligne|Couleur|Transparence|Fermée|Nom du style de tracé")
  * \return Valeur à exporter des paramètres des polylignes fermées
  */

void exportClosedPoly(
    AcDbEntity*& ent,
    AcStringArray& pData );

/**
  * \brief Fonction pour exporter les propriétes des points
  * \param pt L'entité point
  * \param pData liste des informations de chaque point en AcStringArray
  * \pData ("Handle|Type de calque|Type de ligne|Echelle de type de ligne|Matériel|Taille de ligne|Couleur|Transparence|Thickness|Nom du style de tracé|X|Y|Z")
  * \return Valeur à exporter des paramètres des points
  */
void exportPoint(
    AcDbPoint*& pt,
    AcStringArray& pData
);

/**
  * \brief Fonction pour exporter les propriétes des entités
  * \param ent L'entité
  * \param pData liste des informations de chaque entité en AcStringArray
  * \pData ("Handle|Type de calque|Type de ligne|Echelle de type de ligne|Matériel|Taille de ligne|Couleur|Transparence|Thickness|Nom du style de tracé|X|Y|Z")
  * \return Valeur à exporter des paramètres des points
  */
void exportEntity(
    AcDbEntity*& ent,
    AcStringArray& pData );

/**
  * \brief Fonction pour exporter les propriétes des curves
  * \param ent L'entité
  * \param pData liste des informations de chaque curve en AcStringArray
  * \pData ("Handle|Type de calque|Type de ligne|Echelle de type de ligne|Matériel|Taille de ligne|Couleur|Transparence|Thickness|Nom du style de tracé|X|Y|Z")
  * \return Valeur à exporter des paramètres des points
  */
void exportCurve(
    AcDbEntity*& ent,
    AcStringArray& pData );

/**
  * \brief Fonction pour exporter les propriétes des blocks
  * \param pt L'entité bloc
  * \param pData propriétes de l'entité bloc en AcStringArray
  * \pData ("Handle|Nom du bloc|X|Y|Z|Type de calque|Type de ligne|Echelle de type de ligne|Matériel|Taille de ligne|Couleur|Transparence|Nom de style de trace")
  * \return Valeur à exporter des paramètres des blocs
  */
void exportBlock(
    AcDbBlockReference*& blk,
    AcStringArray& pData
);

/**
  * \brief Fonction pour exporter les propriétes des textes
  * \param pt L'entité texte
  * \param pData propriétes de l'entité texte en AcStringArray
  * \pData ("Handle|X|Y|Z|Align_X|Align_Y|Align_Z|Couleur|Type de calque|Materiaux|Nom de style de trace|Thickness|Align_Vertical|Align_Horizontal|Style|Contenu|Taille|Facteur de largeur|Rotation|Oblique")
  * \return Valeur à exporter des paramètres des textes
  */
void exportText(
    AcDbText*& txt,
    AcStringArray& pData
);

/**
  * \brief Fonction pour exporter les propriétes des textes multiples
  * \param pt L'entité texte multiple
  * \param pData propriétes de l'entité texte multiple en AcStringArray
  * \return Valeur à exporter des paramètres des textes multiple
  */
void exportMText(
    AcDbMText*& mTxt,
    AcStringArray& pData
);

bool writeXmlHeader( ofstream& xmlPath );

bool writeXmlFooter( ofstream& xmlPath );

vector<LinePoly> poly3dToSeg( AcDb3dPolyline* poly3d );

bool exportPoly3dToXml( ofstream& xmlPath,
    const vector<LinePoly>& segPoly,
    const AcString& layerPoly,
    const int& laySize,
    const int& i );

int exportFaceToXml(
    ofstream* file,
    string& sLayer );

void exportFaceToXml(
    ofstream* file,
    const vector<FacesInLayer>& vecFl );

//Export lines pour la commande EXPORTLINE
void exportLines(
    AcDbLine*& _line,
    AcStringArray& _lData );

//Export lines pour la commande EXPORTCIRCLE
void exportCircles(
    AcDbCircle*& _circle,
    AcStringArray& _cData );

//Export hatcht pour la commande EXPORTHATCH
void exportHatch(
    AcDbHatch*& _hatch,
    AcStringArray& _hData );

void writePrjOrNot( const AcString& filePath );

int createObjectLine( AcDbEntity*, double*&, double*&, double*& );
