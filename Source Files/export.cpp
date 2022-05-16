#pragma once
#include "export.h"
#include "poly2DEntity.h"
#include "poly3DEntity.h"
#include "blockEntity.h"
#include "textEntity.h"
#include "file.h"
#include "print.h"
#include <set>
#include <iostream>
#include <iomanip>
#include "libxl.h"
#include <unordered_map>

#include <geell2d.h>
#include <genurb2d.h>
#include <hatchEntity.h>
#include <boundingBox.h>
#include <circleEntity.h>
#include "xlnt/xlnt.hpp"

using namespace libxl;

/* ==> Utils function shape <== */

void getFieldsInfo( const DBFHandle& hDBF,
    int( &attributesId )[NB_ATT],
    const AcString& configPath,
    std::vector<string>& defaultValues )
{
    // Initialisation des paramètres
    char pszFieldName[12];
    int pnWidth = 0;
    int pnDecimals = 0;
    
    /// 1. On met à jour attributesId directement avec le nom des champs QGIS
    
    // Récupération du nombre de champs définis dans la table
    int nbFields = DBFGetFieldCount( hDBF );
    
    // Boucle sur tous les champs du shapefile
    for( int i = 0; i < nbFields; i++ )
    {
        // On récupère le nom de l'attribut i du shapefile
        DBFGetFieldInfo( hDBF, i, pszFieldName, &pnWidth, &pnDecimals );
        
        // Passage en majuscules pour gérer les différents cas
        std::string attName( pszFieldName );
        std::transform( attName.begin(), attName.end(), attName.begin(), ::toupper );
        
        // Gestion de chaque attribut
        for( int j = 0; j < NB_ATT; j++ )
        {
            // On compare le nom du champ avec tous les noms possibles de l'attribut
            for( int k = 0; k < 5; k++ )
            {
                std::string opt = attOptions[j][k];
                
                // Si une version du nom correspond au champ, on garde en mémoire l'indice de la colonne
                if( opt != "" && opt == attName )
                {
                    attributesId[j] = i;
                    break;
                }
            }
        }
    }
    
    /// 2. (Opt) On met à jour attributesId avec le fichier de config
    
    // Ouverture du fichier de config
    FileReader* fr = new FileReader( configPath, _T( "\t" ) );
    
    // Lien entre attributesId et le fichier de config codé en dur
    int const configIndices[10] = { LAYER, COLOR, LINETYPE, LINESCALE, LINEWEIGHT, -1, -1, -1, CONSTANTWIDTH, HATCH };
    
    for( int i = 0; i < NB_ATT; i++ )
        defaultValues.push_back( "" );
        
    // On parcourt les lignes de la config qui concernent les paramètres de dessin
    for( int i = 1; i < 10; i++ )
    {
        // On récupère la valeur du champ associé au paramètre
        string field;
        fr->getValue( i, 2, field );
        AcString acfield = strToAcStr( field );
        acutPrintf( acfield );
        
        // On récupère la valeur par défaut
        string defaultValue;
        fr->getValue( i, 1, defaultValue );
        
        if( configIndices[i - 1] != -1 )
        {
            // On ajoute la valeur par défaut dans le vecteur
            defaultValues[configIndices[i - 1]] = defaultValue;
        }
        
        // On cherche son index dans la table QGIS
        int fieldIndex = DBFGetFieldIndex( hDBF, field.c_str() );
        acutPrintf( _T( "%i" ), fieldIndex );
        
        // Si ce champ existe, on met à jour attributesId
        if( fieldIndex != -1 )
        {
            acutPrintf( _T( "%i" ), fieldIndex );
            
            if( configIndices[i] != -1 )
                attributesId[configIndices[i - 1]] = fieldIndex;
        }
    }
}

void getAttFieldIndex( const DBFHandle& hDBF, int( &attributesId )[NB_ATT] )
{
    // Initialisation des paramètres
    char pszFieldName[12];
    int pnWidth = 0;
    int pnDecimals = 0;
    
    // Récupération du nombre de champs définis dans la table
    int nbFields = DBFGetFieldCount( hDBF );
    
    // Boucle sur tous les champs du shapefile
    for( int i = 0; i < nbFields; i++ )
    {
        // On récupère le nom de l'attribut i du shapefile
        DBFGetFieldInfo( hDBF, i, pszFieldName, &pnWidth, &pnDecimals );
        
        // Passage en majuscules pour gérer les différents cas
        std::string attName( pszFieldName );
        std::transform( attName.begin(), attName.end(), attName.begin(), ::toupper );
        
        // Gestion de chaque attribut
        for( int j = 0; j < NB_ATT; j++ )
        {
            // On compare le nom du champ avec tous les noms possibles de l'attribut
            for( int k = 0; k < 5; k++ )
            {
                std::string opt = attOptions[j][k];
                
                // Si une version du nom correspond au champ, on garde en mémoire l'indice de la colonne
                if( opt != "" && opt == attName )
                {
                    attributesId[j] = i;
                    break;
                }
            }
        }
    }
}

Acad::ErrorStatus setLayer( const DBFHandle& hDBF,
    AcDbEntity* pEnt,
    const int& i,
    const int& layerFieldIndex )
{
    Acad::ErrorStatus es;
    
    // Récupération du calque courant
    AcDbObjectId currentLayerId = getCurrentLayerId();
    
    // S'il y a un attribut calque dans le shape
    if( layerFieldIndex != -1 )
    {
        // On récupère le calque de la shape
        // Si le champ est null, on set avec le calque courant
        if( DBFIsAttributeNULL( hDBF, i, layerFieldIndex ) )
        {
            if( es = pEnt->setLayer( currentLayerId ) )
            {
                pEnt->close();
                return es;
            }
            
            return Acad::eOk;
        }
        
        // Si le champ n'est pas null, on récupère le nom du calque
        else
        {
            string layer = DBFReadStringAttribute( hDBF, i, layerFieldIndex );
            
            // On crée le calque si le calque n'existe pas
            if( createLayer( strToAcStr( layer ) ) != Acad::eOk )
            {
                _T( "\n Problème de calque" );
                pEnt->close();
                return Acad::eBadLayerName;
            }
            
            // On met l'entité dans ce calque
            return pEnt->setLayer( strToAcStr( layer ) );
        }
    }
    
    // Si la table n'a pas de champ pour le calque, on met le calque courant
    else
    {
        if( es = pEnt->setLayer( currentLayerId ) )
        {
            pEnt->close();
            return es;
        }
    }
    
    return Acad::eOk;
}


Acad::ErrorStatus setColor( const DBFHandle& hDBF,
    AcDbEntity* pEnt,
    const int& i,
    const int& colorFieldIndex )
{
    Acad::ErrorStatus es = Acad::eOk;
    
    int width = 0;
    int decimals = 0;
    char pszFieldName[12];
    
    AcCmColor acadColor = AcCmColor();
    
    if( colorFieldIndex != -1 )
    {
        // On récupère le type de l'attribut: entier, string ...
        DBFFieldType type = DBFGetFieldInfo( hDBF,
                colorFieldIndex,
                pszFieldName,
                &width,
                &decimals );
                
        if( type == FTString )
        {
            string color = DBFReadStringAttribute( hDBF, i, colorFieldIndex );
            
            if( color != "" )
            {
                std::vector<string> rgb = splitString( color, "," );
                
                // On pourrait rendre le résultat plus robuste en testant si r g et b sont bien des int
                if( rgb.size() != 3 )
                    print( "Warning: RGB invalide. L'objet prend la couleur par défaut." );
                    
                else
                {
                    int r = atoi( rgb[0].c_str() );
                    int g = atoi( rgb[1].c_str() );
                    int b = atoi( rgb[2].c_str() );
                    acadColor.setRGB( r, g, b );
                    es = pEnt->setColor( acadColor );
                    
                    if( es != Acad::eOk )
                    {
                        acutPrintf( _T( "\n Cette valeur de couleur n'est pas disponible" ) );
                        return es;
                    }
                }
            }
        }
        
        else if( type == FTInteger )
        {
            // On récupère la couleur dans le fichier shp
            int color = DBFReadIntegerAttribute( hDBF, i, colorFieldIndex );
            
            // On la set sur l'entité
            acadColor.setColorIndex( color );
            es = pEnt->setColor( acadColor );
            
            if( es != Acad::eOk )
            {
                acutPrintf( _T( "\n Cette valeur de couleur n'est pas disponible" ) );
                return es;
            }
        }
    }
    
    return Acad::eOk;
}

Acad::ErrorStatus setLinetype( const DBFHandle& hDBF,
    AcDbEntity* pEnt,
    const int& i,
    const int& linetypeFieldIndex )
{
    Acad::ErrorStatus es;
    
    // Si la table QGIS a un champ type de ligne
    if( linetypeFieldIndex != -1 )
    {
        // On récupère l'attribut type de ligne
        string linetype = DBFReadStringAttribute( hDBF, i, linetypeFieldIndex );
        
        // Si celui-ci n'est pas NULL
        if( linetype != "" )
        {
            // On le met à jour si le type de ligne demandé existe
            if( es = pEnt->setLinetype( strToAcStr( linetype ) ) )
            {
                acutPrintf( _T( "\nCe type de ligne n'est pas disponible" ) );
                return es;
            }
        }
    }
    
    return Acad::eOk;
}


Acad::ErrorStatus setLinescale( const DBFHandle& hDBF,
    AcDbEntity* pEnt,
    const int& i,
    const int& linescaleFieldIndex )
{
    Acad::ErrorStatus es;
    
    // Si l'échelle de type de ligne est dans la table QGIS
    if( linescaleFieldIndex != -1 )
    {
        // On vérifie que l'attribut n'est pas NULL
        if( !DBFIsAttributeNULL( hDBF, i, linescaleFieldIndex ) )
        {
            // On récupère l'attribut
            double linescale = DBFReadDoubleAttribute( hDBF, i, linescaleFieldIndex );
            
            // On met à jour l'entité
            if( es = pEnt->setLinetypeScale( linescale ) )
            {
                acutPrintf( _T( "\nImpossible de charger l'échelle de type de ligne" ) );
                return es;
            }
        }
    }
    
    return Acad::eOk;
}

Acad::ErrorStatus setLineweight( const DBFHandle& hDBF,
    AcDbEntity* pEnt,
    const int& i,
    const int& lineweightFieldIndex )
{
    Acad::ErrorStatus es = Acad::eOk;
    
    // Si l'épaisseur de type de ligne est dans la table QGIS
    if( lineweightFieldIndex != -1 )
    {
        // On vérifie que l'attribut n'est pas NULL
        if( !DBFIsAttributeNULL( hDBF, i, lineweightFieldIndex ) )
        {
            // On récupère l'attribut
            int lineweight = DBFReadIntegerAttribute( hDBF, i, lineweightFieldIndex );
            
            // On récupère l'enum LineWeight dans la database
            AcDbDatabase* pDB = acdbHostApplicationServices()->workingDatabase();
            AcDb::LineWeight lw = pDB->getNearestLineWeight( lineweight );
            
            if( lw != lineweight )
                acutPrintf( _T( "\nEpaisseur de ligne %i remplacée par %i" ), lineweight, lw );
                
            // On met à jour l'entité
            if( es = pEnt->setLineWeight( lw ) )
            {
                acutPrintf( _T( "\nImpossible de charger l'épaisseur de type de ligne" ) );
                return es;
            }
        }
    }
    
    return Acad::eOk;
}

Acad::ErrorStatus setConstantWidth( const DBFHandle& hDBF,
    AcDbPolyline* poly,
    const int& i,
    const int& constantwidthFieldIndex )
{
    Acad::ErrorStatus es;
    
    // Si la largeur constante est dans la table QGIS
    if( constantwidthFieldIndex != -1 )
    {
        // On vérifie que l'attribut n'est pas NULL
        if( !DBFIsAttributeNULL( hDBF, i, constantwidthFieldIndex ) )
        {
            // On récupère l'attribut
            double width = DBFReadDoubleAttribute( hDBF, i, constantwidthFieldIndex );
            
            // On met à jour l'entité
            if( es = poly->setConstantWidth( width ) )
            {
                acutPrintf( _T( "\nImpossible de charger la largeur constante" ) );
                return es;
            }
        }
    }
    
    return Acad::eOk;
}

Acad::ErrorStatus setHatch3d( const DBFHandle& hDBF,
    vector<AcDb3dPolyline*> poly,
    const int& i,
    const int& hatchFieldIndex,
    int( &attributesId )[NB_ATT] )
{
    Acad::ErrorStatus es;
    
    // Initialisation de la hachure
    AcDbHatch* hatch = new AcDbHatch();
    addToModelSpace( hatch );
    string patternName = "";
    
    // Si le champ hachure est dans la table QGIS
    if( hatchFieldIndex != -1 )
    {
        // On récupère l'attribut hachure
        patternName = DBFReadStringAttribute( hDBF, i, hatchFieldIndex );
    }
    
    // Si celui-ci est NULL, on met la hachure courante
    if( patternName == "" )
    {
        // Récuperer le nom de la hachure courante
        patternName = acdbHostApplicationServices()->workingDatabase()->getHPNAME();
    }
    
    // On met à jour le pattern
    if( es = insert3DHatch( poly, hatch, strToAcStr( patternName ) ) )
    {
        acutPrintf( _T( "\nCe type de hachure n'est pas disponible" ) );
        return es;
    }
    
    if( es = setCharacteristics( hDBF, hatch, i, attributesId ) )
    {
        hatch->close();
        return es;
    }
    
    hatch->close();
    
    return Acad::eOk;
}

Acad::ErrorStatus setHatch2d( const DBFHandle& hDBF,
    vector<AcDbPolyline*> poly,
    const int& i,
    const int& hatchFieldIndex,
    int( &attributesId )[NB_ATT] )
{
    Acad::ErrorStatus es;
    
    // Initialisation de la hachure
    AcDbHatch* hatch = new AcDbHatch();
    addToModelSpace( hatch );
    string patternName = "";
    
    // Si le champ hachure est dans la table QGIS
    if( hatchFieldIndex != -1 )
    {
        // On récupère l'attribut hachure
        patternName = DBFReadStringAttribute( hDBF, i, hatchFieldIndex );
    }
    
    // Si celui-ci est NULL, on met la hachure courante
    if( patternName == "" )
    {
        // Récuperer le nom de la hachure courante
        patternName = acdbHostApplicationServices()->workingDatabase()->getHPNAME();
    }
    
    // On met à jour le pattern
    if( es = insert2DHatch( poly, hatch, strToAcStr( patternName ) ) )
    {
        acutPrintf( _T( "\nCe type de hachure n'est pas disponible" ) );
        return es;
    }
    
    if( es = setCharacteristics( hDBF, hatch, i, attributesId ) )
    {
        hatch->close();
        return es;
    }
    
    hatch->close();
    
    return Acad::eOk;
}

Acad::ErrorStatus setAttributes( AcDbBlockReference* pBlock,
    const DBFHandle& hDBF,
    const int& i,
    AcDbEntity* poly )
{
    int width = 0;
    int decimals = 0;
    AcString attValue = _T( "" );
    
    // On récupère les attributs du bloc
    map< AcString, AcString > attList = getBlockAttWithValuesList( pBlock );
    
    // On boucle sur tous les attributs du bloc
    map< AcString, AcString >::iterator it = attList.begin();
    
    while( it != attList.end() )
    {
        // On récupère le nom de l'attribut
        string attName = acStrToStr( it->first );
        
        // On récupère l'index de l'attribut dans le shape
        int attIndex = DBFGetFieldIndex( hDBF, &attName[0] );
        
        // Si l'attribut existe bien
        if( attIndex != -1 )
        {
            // On récupère le type de l'attribut: entier, string ...
            DBFFieldType type = DBFGetFieldInfo( hDBF,
                    attIndex,
                    &attName[0],
                    &width,
                    &decimals );
                    
            // Attribut = string
            if( type == FTString )
            {
                attValue = strToAcStr( DBFReadStringAttribute( hDBF,
                            i,
                            attIndex ) );
            }
            
            // Attribut = integer
            else if( type == FTInteger )
            {
                attValue = numberToAcString( DBFReadIntegerAttribute( hDBF,
                            i,
                            attIndex ), 0 );
            }
            
            // Attribut = double
            else if( type == FTDouble )
            {
                attValue = numberToAcString( DBFReadDoubleAttribute( hDBF,
                            i,
                            attIndex ),
                        decimals );
            }
            
            // Maintenant on va setter la valeur dans l'attribut du block
            // Si la valeur de l'attribut est renseignée dans le shape
            if( attValue != _T( "" ) )
            {
                if( !setAttributValue( pBlock,
                        strToAcStr( toUpperString( attName ) ),
                        attValue ) )
                    acutPrintf( _T( "\n Problème avec l'attribut" ) + strToAcStr( attName ) );
            }
            
            // Si la valeur de l'attribut est vide dans le shape
            // Alors on le remplit avec sa valeur par défaut
            else
            {
                AcDbAttribute* att = getAttributObject( pBlock,
                        strToAcStr( toUpperString( attName ) ),
                        AcDb::kForWrite );
                        
                if( !setAttributValue( pBlock,
                        strToAcStr( toUpperString( attName ) ),
                        att->textString() ) )
                    acutPrintf( _T( "\n Problème avec l'attribut" ) + strToAcStr( attName ) );
                    
                att->close();
            }
        }
        
        // On passe à l'attribut suivant
        it++;
    }
    
    // Si le bloc a un attribut Handle alors on va mettre celui de la polyligne à l'intérieur
    AcDbHandle handle;
    
    if( poly )
        handle = poly->handle();
        
    else
        return Acad::eOk;
        
    if( handle.isNull() )
        return Acad::eNullHandle;
        
    // On récupère le Handle en string
    setAttributValue( pBlock,
        _T( "HANDLE" ),
        handle );
        
    return Acad::eOk;
}

Acad::ErrorStatus setProperties( AcDbBlockReference*& pBlock,
    const DBFHandle& hDBF,
    const int& i,
    AcDbEntity* poly )
{
    Acad::ErrorStatus es = Acad::eOk;
    
    // Ouverture du bloc dynamique
    AcDbDynBlockReference* pDynBlock = new AcDbDynBlockReference( pBlock );
    AcDbDynBlockReferencePropertyArray properties;
    
    // Rcupration des proprits
    pDynBlock->getBlockProperties( properties );
    
    // Initialisation des paramtres
    AcDbDynBlockReferenceProperty blkProp;
    int width = 0;
    int decimals = 0;
    char pszFieldName[12];
    
    // Itration sur les proprits
    int nbProperties = properties.length();
    
    for( int ii = 0; ii < nbProperties; ii++ )
    {
        blkProp = properties[ii];
        AcString propName = blkProp.propertyName();
        
        // On récupère l'index de l'attribut dans le shape
        int propIndex = DBFGetFieldIndex( hDBF, propName );
        
        // Si une propriété apparaît
        if( propIndex != -1 )
        {
            // On vérifie que le champ n'est pas nul
            if( !DBFIsAttributeNULL( hDBF, i, propIndex ) )
            {
                AcDbEvalVariant eval;
                DBFFieldType type = DBFGetFieldInfo( hDBF,
                        propIndex,
                        pszFieldName,
                        &width,
                        &decimals );
                        
                string test = DBFReadStringAttribute( hDBF,
                        i,
                        propIndex );
                        
                // Attribut = string
                if( type == FTString )
                {
                    eval = AcDbEvalVariant( strToAcStr( DBFReadStringAttribute( hDBF,
                                    i,
                                    propIndex ) ) );
                }
                
                // Attribut = integer
                else if( type == FTInteger )
                {
                    eval = AcDbEvalVariant( Adesk::Int32( DBFReadIntegerAttribute( hDBF,
                                    i,
                                    propIndex ) ) );
                }
                
                
                // Attribut = double
                else if( type == FTDouble )
                {
                    eval = AcDbEvalVariant( DBFReadDoubleAttribute( hDBF,
                                i,
                                propIndex ) );
                }
                
                // Mise à jour de la propriété
                if( es = blkProp.setValue( eval ) )
                {
                    delete pDynBlock;
                    return es;
                }
            }
        }
    }
    
    return es;
}

Acad::ErrorStatus setCharacteristics( const DBFHandle& hDBF,
    AcDbEntity* pEnt,
    const int& i,
    int( &attributesId )[NB_ATT] )
{
    Acad::ErrorStatus es;
    
    // On met le point dans le bon calque
    if( es = setLayer( hDBF,
                pEnt,
                i,
                attributesId[LAYER] ) )
    {
        pEnt->close();
        return es;
    }
    
    // On met le point dans la bonne couleur
    if( es = setColor( hDBF,
                pEnt,
                i,
                attributesId[COLOR] ) )
    {
        pEnt->close();
        return es;
    }
    
    // On donne le bon type de ligne
    if( es = setLinetype( hDBF,
                pEnt,
                i,
                attributesId[LINETYPE] ) )
    {
        pEnt->close();
        return es;
    }
    
    // Echelle de type de ligne
    if( es = setLinescale( hDBF,
                pEnt,
                i,
                attributesId[LINESCALE] ) )
    {
        pEnt->close();
        return es;
    }
    
    // Epaisseur de type de ligne
    if( es = setLineweight( hDBF,
                pEnt,
                i,
                attributesId[LINEWEIGHT] ) )
    {
        pEnt->close();
        return es;
    }
    
    return es;
}

Acad::ErrorStatus askBlockName( const SHPHandle& hSHP,
    const DBFHandle& hDBF,
    const int& blockFieldIndex,
    AcString& blockName )
{
    // On propose par défaut le nom du dernier bloc inséré
    AcString insname;//= getStringVariable(_T("INSNAME"));
    
    // Si l'attribut bloc n'est pas défini
    // On demande au dessinateur quel bloc il veut insérer
    // S'il rentre une valeur, on l'enregistre pour la reproposer plus tard
    AcString message = _T( "" );
    
    if( insname == AcString::kEmpty )
        message = _T( "\nEntrer le nom du bloc à insérer : [<Aucun>/Taper le nom du bloc]\n" );
    else
        message = _T( "\nEntrer le nom du bloc à insérer: [<" ) + insname + _T( ">/Aucun/Taper le nom du bloc]\n" );
        
    // Si le nom du bloc n'est pas en attribut du shape
    if( blockFieldIndex == -1 )
    {
        // On demande de saisir le nom du bloc
        int res = askString(
                acStrToAcharPointeur( message ),
                blockName,
                acStrToAcharPointeur( insname )
            );
            
        // Si le user fait juste ENTREE
        if( res == RTNORM && blockName == AcString::kEmpty )
        {
            print( blockName );
            blockName = insname;
            return Acad::eOk;
        }
        
        else if( res == RTNORM && areEqualNoCase( blockName, _T( "Aucun" ) ) )
        {
            print( "Aucun" );
            blockName = _T( "" );
            return Acad::eOk;
        }
        
        // Si le user fait ECHAP
        else if( res == RTCAN )
        {
            // On ferme le fichier shp
            SHPClose( hSHP );
            
            // Fermer le fichier d'attributs
            DBFClose( hDBF );
            
            // On quitte
            return Acad::eNotApplicable;
        }
        
        // Si le user entre une valeur au CLAVIER
        else
            return Acad::eOk;
    }
    
    else return Acad::eOk;
}


Acad::ErrorStatus insertBlock( const DBFHandle& hDBF,
    SHPObject* iShape,
    const int& i,
    const AcGePoint3d& ptInsert,
    const AcString& blockName,
    int( &attributesId )[NB_ATT],
    AcDbEntity* poly )
{
    Acad::ErrorStatus es;
    
    // On vérifie que le bloc existe
    //Recuperer l'identifiant de la definition de bloc
    AcDbObjectId blockId = getBlockDefId( blockName );
    
    if( blockId.isNull() )
    {
        print( "Le bloc %s n'est pas défini", blockName );
        return Acad::eOk;
    }
    
    // Avant de créer le bloc, on récupère les informations de rotation et d'échelle pour les appliquer dès l'insertion du bloc
    double angle = 0.0;
    double scaleX = 1.0;
    double scaleY = 1.0;
    double scaleZ = 1.0;
    getAngleAndScale( hDBF, i, attributesId, angle, scaleX, scaleY, scaleZ );
    
    AcDbBlockReference* pBlock = insertBlockReference( blockName, ptInsert, angle, scaleX, scaleY, scaleZ );
    
    // On set les caractéristiques du bloc
    if( es = setCharacteristics( hDBF, pBlock, i, attributesId ) )
    {
        pBlock->close();
        return es;
    }
    
    // On set tous les autres attributs du bloc
    if( es = setAttributes( pBlock, hDBF, i, poly ) )
    {
        pBlock->close();
        return es;
    }
    
    // Si le bloc est dynamique, on set aussi ses propriétés
    AcDbDynBlockTableRecord pDynBlockDef( pBlock->blockTableRecord() );
    
    if( pDynBlockDef.isDynamicBlock() )
    {
        if( es = setProperties( pBlock, hDBF, i, poly ) )
        {
            pBlock->close();
            //return es;
        }
    }
    
    // On ferme le bloc
    return pBlock->close();
}


Acad::ErrorStatus importShp( const SHPHandle& hSHP,
    const DBFHandle& hDBF,
    const int& pnShapeType,
    const int& pnEntities,
    int( &attributesId )[NB_ATT],
    AcString blockName,
    bool drawHatch )
{
    if( blockName != _T( "" ) )
    {
        // On insère le block
        AcDbBlockTableRecord* pBlkDef = getBlockDef( blockName );
        
        // Si la définition du bloc n'est pas dans le dessin
        // On sort de la fonction
        if( !pBlkDef )
        {
            acutPrintf( _T( "\nLe bloc " ) + blockName + _T( " n'est pas défini" ) );
            blockName = _T( "" );
        }
        
        else
            pBlkDef->close();
    }
    
    int blockFieldIndex = attributesId[BLOCK];
    
    // En fonction du type de shape, on appelle la bonne fonction d'import
    if( pnShapeType == SHPT_POINT ||
        pnShapeType == SHPT_POINTZ ||
        pnShapeType == SHPT_POINTM ||
        pnShapeType == SHPT_MULTIPOINT ||
        pnShapeType == SHPT_MULTIPOINTZ ||
        pnShapeType == SHPT_MULTIPOINTM )
        return importPoints( hSHP,
                hDBF,
                pnEntities,
                attributesId,
                blockFieldIndex,
                blockName );
                
    else if( pnShapeType == SHPT_ARC )
        return importAcDbPolyline( hSHP,
                hDBF,
                pnEntities,
                attributesId,
                blockFieldIndex,
                blockName );
                
    else if( pnShapeType == SHPT_ARCZ ||
        pnShapeType == SHPT_ARCM )
        return importAcDbPolyline3d( hSHP,
                hDBF,
                pnEntities,
                attributesId,
                blockFieldIndex,
                blockName );
                
    else if( pnShapeType == SHPT_POLYGON )
        return importPolygons2d( hSHP,
                hDBF,
                pnEntities,
                attributesId,
                blockFieldIndex,
                blockName,
                drawHatch );
                
    else if( pnShapeType == SHPT_POLYGONZ ||
        pnShapeType == SHPT_POLYGONM )
        return importPolygons3d( hSHP,
                hDBF,
                pnEntities,
                attributesId,
                blockFieldIndex,
                blockName,
                drawHatch );
                
    else
    {
        acutPrintf( _T( "\nEntités dans le fichier .shp non reconnues" ) );
        return Acad::eInvalidInput;
    }
}


Acad::ErrorStatus importPoints( const SHPHandle& hSHP,
    const DBFHandle& hDBF,
    const int& pnEntities,
    int( &attributesId )[NB_ATT],
    const int& blockFieldIndex,
    AcString blockName )
{
    Acad::ErrorStatus es;
    int nbPoints = 0;
    
    // On crée la progressBar
    ProgressBar prog = ProgressBar( _T( "Progression de l'import:" ), pnEntities );
    
    // On boucle sur le nombre d'objets
    for( int i = 0; i < pnEntities; i++ )
    {
        // Shape courant
        SHPObject* iShape = SHPReadObject( hSHP, i );
        int nbVertices = iShape->nVertices;
        
        for( int j = 0; j < nbVertices; ++j )
        {
            if( blockFieldIndex != -1 )
                blockName = strToAcStr( DBFReadStringAttribute( hDBF, i, blockFieldIndex ) );
                
            // Coordonnées de l'objet
            double x = iShape->padfX[j];
            double y = iShape->padfY[j];
            double z = iShape->padfZ[j];
            
            AcGePoint3d ptInsert = AcGePoint3d( x, y, z );
            
            // Si on a un nom de bloc valide, on l'insère à la place des points
            if( blockName != _T( "" ) )
            {
                if( es = insertBlock( hDBF,
                            iShape,
                            i,
                            ptInsert,
                            blockName,
                            attributesId ) )
                {
                    SHPDestroyObject( iShape );
                    return es;
                }
            }
            
            // Sinon on insère un AcDbPoint
            else
            {
                AcDbPoint* point = new AcDbPoint( ptInsert );
                
                // On set les caractéristiques
                if( es = setCharacteristics( hDBF, point, i, attributesId ) )
                {
                    delete point;
                    return es;
                }
                
                // On ajoute le point à la Db
                addToModelSpace( point );
                
                // On ferme le point
                point->close();
            }
            
            nbPoints++;
        }
        
        // On supprime la shape
        SHPDestroyObject( iShape );
        
        // On avance la barre de progression
        prog.moveUp( i );
    }
    
    print( "Points insérés: %i", nbPoints );
    
    // Si le bloc entré par l'utilisateur existe, on change le insname après l'insertion
    if( getBlockDefId( blockName ) != AcDbObjectId::kNull )
        setStringVariable( _T( "INSNAME" ), blockName );
        
    return Acad::eOk;
}


Acad::ErrorStatus importAcDbPolyline3d( const SHPHandle& hSHP,
    const DBFHandle& hDBF,
    const int& pnEntities,
    int( &attributesId )[NB_ATT],
    const int& blockFieldIndex,
    AcString blockName )
{
    Acad::ErrorStatus es;
    int nbPolylines3d = 0;
    
    // On crée la progressBar
    ProgressBar prog = ProgressBar( _T( "Progression de l'import:" ), pnEntities );
    
    AcGePoint3d ptInsert = AcGePoint3d::kOrigin;
    
    // On boucle sur le nombre d'objets
    for( int i = 0; i < pnEntities; i++ )
    {
        // Shape courant
        SHPObject* iShape = SHPReadObject( hSHP, i );
        
        // On doit suivre le compte des vertex dans le cas où la shape contient plusieurs entités
        int nbVertices = 0;
        
        int nbParts = iShape->nParts;
        
        // Dans le cas iShape NULL du coup on continue la boucle
        if( nbParts == 0 )
        {
            SHPDestroyObject( iShape );
            continue;
        }
        
        // On itère sur le nombre d'entités
        for( int j = 0; j < nbParts; ++j )
        {
            // On crée la poly
            AcDb3dPolyline* poly = new AcDb3dPolyline();
            
            // On l'ajoute à la Db
            addToModelSpace( poly );
            
            // Nombre de Sommets de la polyligne
            int nbPartVertices;
            
            if( j < nbParts - 1 )
                nbPartVertices = iShape->panPartStart[j + 1] - iShape->panPartStart[j];
            else
                nbPartVertices = iShape->nVertices - iShape->panPartStart[nbParts - 1];
                
            // Boucle sur les vertex
            for( int k = 0; k < nbPartVertices; k++ )
            {
                // On crée le vertex 3d
                AcDb3dPolylineVertex* vertex = new AcDb3dPolylineVertex( AcGePoint3d( iShape->padfX[nbVertices + k], iShape->padfY[nbVertices + k], iShape->padfZ[nbVertices + k] ) );
                
                // On le rajoute à la poly
                poly->appendVertex( vertex );
                
                // On le ferme
                vertex->close();
            }
            
            // On actualise le nombre de sommets
            nbVertices += nbPartVertices;
            
            // On set les caractéristiques
            if( es = setCharacteristics( hDBF, poly, i, attributesId ) )
            {
                delete poly;
                return es;
            }
            
            // On lit le nom du bloc à insérer
            if( blockFieldIndex != -1 )
                blockName = strToAcStr( DBFReadStringAttribute( hDBF, i, blockFieldIndex ) );
                
            // Si on a un nom de bloc valide, on l'insère au milieu de la polyligne
            if( blockName != _T( "" ) )
            {
                // On récupère le milieu de la poly
                if( getMilieu( poly, ptInsert ) != Acad::eOk )
                {
                    poly->close();
                    continue;
                }
                
                // On insère un bloc au mileu de la polyligne
                if( insertBlock( hDBF,
                        iShape,
                        i,
                        ptInsert,
                        blockName,
                        attributesId,
                        poly ) != Acad::eOk )
                {
                    SHPDestroyObject( iShape );
                    poly->close();
                    continue;
                }
            }
            
            // On ferme la poly
            poly->close();
            
            nbPolylines3d++;
        }
        
        // On supprime la shape
        SHPDestroyObject( iShape );
        
        // On avance la barre de progression
        prog.moveUp( i );
    }
    
    print( "Polylignes 3d insérées: %d", nbPolylines3d );
    
    // Si le bloc entré par l'utilisateur existe, on change le insname après l'insertion
    if( getBlockDefId( blockName ) != AcDbObjectId::kNull )
        setStringVariable( _T( "INSNAME" ), blockName );
        
    return Acad::eOk;
}


Acad::ErrorStatus importAcDbPolyline( const SHPHandle& hSHP,
    const DBFHandle& hDBF,
    const int& pnEntities,
    int( &attributesId )[NB_ATT],
    const int& blockFieldIndex,
    AcString blockName )
{
    Acad::ErrorStatus es;
    int nbPolylines = 0;
    
    // On crée la progressBar
    ProgressBar prog = ProgressBar( _T( "Progression de l'import:" ), pnEntities );
    
    AcGePoint3d ptInsert = AcGePoint3d::kOrigin;
    
    // On boucle sur le nombre d'objets
    for( int i = 0; i < pnEntities; i++ )
    {
        // Shape courant
        SHPObject* iShape = SHPReadObject( hSHP, i );
        
        // On doit suivre le compte des vertex dans le cas où la shape contient plusieurs entités
        int nbVertices = 0;
        
        int nbParts = iShape->nParts;
        
        // Dans le cas iShape NULL du coup on continue la boucle
        if( nbParts == 0 )
        {
            SHPDestroyObject( iShape );
            continue;
        }
        
        // On itère sur le nombre d'entités
        for( int j = 0; j < nbParts; ++j )
        {
            // On crée la poly
            AcDbPolyline* poly = new AcDbPolyline();
            
            // On l'ajoute à la Db
            addToModelSpace( poly );
            
            // Nombre de Sommets de la polyligne
            int nbPartVertices;
            
            if( j < nbParts - 1 )
                nbPartVertices = iShape->panPartStart[j + 1] - iShape->panPartStart[j];
            else
                nbPartVertices = iShape->nVertices - iShape->panPartStart[nbParts - 1];
                
            // Boucle sur les vertex
            for( int k = 0; k < nbPartVertices; k++ )
            {
                // On le rajoute à la poly
                poly->addVertexAt( k, AcGePoint2d( iShape->padfX[nbVertices + k], iShape->padfY[nbVertices + k] ) );
            }
            
            // On actualise le nombre de sommets
            nbVertices += nbPartVertices;
            
            // On set les caractéristiques
            if( es = setCharacteristics( hDBF, poly, i, attributesId ) )
            {
                delete poly;
                return es;
            }
            
            // On set la largeur constante
            if( es = setConstantWidth( hDBF, poly, i, attributesId[CONSTANTWIDTH] ) )
            {
                delete poly;
                return es;
            }
            
            // On lit le nom du bloc à insérer
            if( blockFieldIndex != -1 )
                blockName = strToAcStr( DBFReadStringAttribute( hDBF, i, blockFieldIndex ) );
                
            // Si on a un nom de bloc valide, on l'insère au milieu de la polyligne
            if( blockName != _T( "" ) )
            {
                // On récupère le milieu de la poly
                if( getMilieu( poly, ptInsert ) != Acad::eOk )
                {
                    poly->close();
                    continue;
                }
                
                // On insère un bloc au mileu de la polyligne
                if( insertBlock( hDBF,
                        iShape,
                        i,
                        ptInsert,
                        blockName,
                        attributesId,
                        poly ) != Acad::eOk )
                {
                    SHPDestroyObject( iShape );
                    poly->close();
                    continue;
                }
            }
            
            // On ferme la poly
            poly->close();
            
            nbPolylines++;
        }
        
        // On supprime la shape
        SHPDestroyObject( iShape );
        
        // On avance la barre de progression
        prog.moveUp( i );
    }
    
    print( "Polylignes 2d insérées: %i", nbPolylines );
    
    // Si le bloc entré par l'utilisateur existe, on change le insname après l'insertion
    if( getBlockDefId( blockName ) != AcDbObjectId::kNull )
        setStringVariable( _T( "INSNAME" ), blockName );
        
    return Acad::eOk;
}


Acad::ErrorStatus importPolygons3d( const SHPHandle& hSHP,
    const DBFHandle& hDBF,
    const int& pnEntities,
    int( &attributesId )[NB_ATT],
    const int& blockFieldIndex,
    AcString blockName,
    bool drawHatch )
{
    Acad::ErrorStatus es;
    int nbPolygons = 0;
    
    // On crée la progressBar
    ProgressBar prog = ProgressBar( _T( "Progression de l'import:" ), pnEntities );
    
    AcGePoint3d ptInsert = AcGePoint3d::kOrigin;
    
    // On boucle sur le nombre d'objets
    for( int i = 0; i < pnEntities; i++ )
    {
        // Shape courant
        SHPObject* iShape = SHPReadObject( hSHP, i );
        
        // Dans le cas où des polygones seraient imbriqués, on doit garder en mémoire les éléments suivants :
        vector<AcDbExtents> boundingBoxes;
        vector<AcDb3dPolyline*> polylines;
        vector<AcGePoint3d> centroids;
        
        // On doit suivre le compte des vertex dans le cas où la shape contient plusieurs entités
        int nbVertices = 0;
        
        int nbParts = iShape->nParts;
        
        // Dans le cas iShape NULL du coup on continue la boucle
        if( nbParts == 0 )
        {
            SHPDestroyObject( iShape );
            continue;
        }
        
        // On itère sur le nombre d'entités
        for( int j = 0; j < nbParts; ++j )
        {
            // On crée la poly
            AcDb3dPolyline* poly = new AcDb3dPolyline();
            
            // On l'ajoute à la Db
            addToModelSpace( poly );
            
            // Nombre de Sommets de la polyligne
            int nbPartVertices;
            
            if( j < nbParts - 1 )
                nbPartVertices = iShape->panPartStart[j + 1] - iShape->panPartStart[j];
            else
                nbPartVertices = iShape->nVertices - iShape->panPartStart[nbParts - 1];
                
            // Boucle sur les vertex
            for( int k = 0; k < nbPartVertices; k++ )
            {
                // On crée le vertex 3d
                AcDb3dPolylineVertex* vertex = new AcDb3dPolylineVertex( AcGePoint3d( iShape->padfX[nbVertices + k], iShape->padfY[nbVertices + k], iShape->padfZ[nbVertices + k] ) );
                
                // On le rajoute à la poly
                poly->appendVertex( vertex );
                
                // On le ferme
                vertex->close();
            }
            
            // On actualise le nombre de sommets
            nbVertices += nbPartVertices;
            
            // On set les caractéristiques de la polyligne
            if( es = setCharacteristics( hDBF, poly, i, attributesId ) )
            {
                delete poly;
                return es;
            }
            
            // On lit le nom du bloc à insérer
            if( blockFieldIndex != -1 )
                blockName = strToAcStr( DBFReadStringAttribute( hDBF, i, blockFieldIndex ) );
                
            // On ferme la poly
            poly->makeClosed();
            
            // On récupère le centroïde de la poly
            if( getCentroid( poly, ptInsert ) != Acad::eOk )
            {
                SHPDestroyObject( iShape );
                
                for( int k = 0; k < polylines.size(); k++ )
                    polylines[k]->close();
                    
                continue;
            }
            
            // On calcule la bounding box
            AcDbExtents bB;
            poly->getGeomExtents( bB );
            
            // On regarde si le polygone courant est contenu ou contient un des polygones déjà examinés
            int index, isContained;
            
            if( isContainedBy( boundingBoxes, isContained, index, bB ) )
            {
                if( blockName != _T( "" ) )
                {
                    AcGePoint3d pt;
                    
                    // On conserve le bloc à insérer pour le plus grand polygone
                    if( isContained == 1 )
                        pt = centroids[index];
                    else
                        pt = ptInsert;
                        
                    if( insertBlock( hDBF, iShape, i, pt, blockName, attributesId ) )
                    {
                        SHPDestroyObject( iShape );
                        
                        for( int k = 0; k < polylines.size(); k++ )
                            polylines[k]->close();
                            
                        continue;
                    }
                }
                
                if( drawHatch )
                {
                    vector<AcDb3dPolyline*> borders;
                    borders.push_back( polylines[index] );
                    borders.push_back( poly );
                    
                    if( es = setHatch3d( hDBF, borders, i, attributesId[HATCH], attributesId ) )
                    {
                        SHPDestroyObject( iShape );
                        
                        for( int k = 0; k < polylines.size(); k++ )
                            polylines[k]->close();
                            
                        continue;
                    }
                }
                
                // On ferme/supprime les objets qui ont été ajoutés des vecteurs de stockage
                polylines[index]->close();
                poly->close();
                polylines.erase( polylines.begin() + index );
                centroids.erase( centroids.begin() + index );
                boundingBoxes.erase( boundingBoxes.begin() + index );
                
                nbPolygons++;
            }
            
            // Sinon, on ajoute les informations aux vecteurs de stockage
            else
            {
                polylines.push_back( poly );
                centroids.push_back( ptInsert );
                boundingBoxes.push_back( bB );
            }
        }
        
        // Il reste à parcourir les entités qui n'ont pas encore été traitées
        for( int j = 0; j < polylines.size(); j++ )
        {
            if( blockName != _T( "" ) )
            {
                if( insertBlock( hDBF, iShape, i, centroids[j], blockName, attributesId ) )
                {
                    SHPDestroyObject( iShape );
                    
                    for( int k = 0; k < polylines.size(); k++ )
                        polylines[k]->close();
                        
                    continue;
                }
            }
            
            if( drawHatch )
            {
                vector<AcDb3dPolyline*> borders;
                borders.push_back( polylines[j] );
                
                if( es = setHatch3d( hDBF, borders, i, attributesId[HATCH], attributesId ) )
                {
                    SHPDestroyObject( iShape );
                    
                    for( int k = 0; k < polylines.size(); k++ )
                        polylines[k]->close();
                        
                    continue;
                }
            }
            
            polylines[j]->close();
            nbPolygons++;
        }
        
        // On supprime la shape
        SHPDestroyObject( iShape );
        
        // On avance la barre de progression
        prog.moveUp( i );
    }
    
    print( "Polygones insérés : %i", nbPolygons );
    
    // Si le bloc entré par l'utilisateur existe, on change le insname après l'insertion
    if( getBlockDefId( blockName ) != AcDbObjectId::kNull )
        setStringVariable( _T( "INSNAME" ), blockName );
        
    return Acad::eOk;
}

Acad::ErrorStatus importPolygons2d( const SHPHandle& hSHP,
    const DBFHandle& hDBF,
    const int& pnEntities,
    int( &attributesId )[NB_ATT],
    const int& blockFieldIndex,
    AcString blockName,
    bool drawHatch )
{
    Acad::ErrorStatus es;
    int nbPolygons = 0;
    
    // On crée la progressBar
    ProgressBar prog = ProgressBar( _T( "Progression de l'import:" ), pnEntities );
    
    AcGePoint3d ptInsert = AcGePoint3d::kOrigin;
    
    // On boucle sur le nombre d'objets
    for( int i = 0; i < pnEntities; i++ )
    {
        // Shape courant
        SHPObject* iShape = SHPReadObject( hSHP, i );
        
        // Dans le cas où des polygones seraient imbriqués, on doit garder en mémoire les éléments suivants :
        vector<AcDbExtents> boundingBoxes;
        vector<AcDbPolyline*> polylines;
        vector<AcGePoint3d> centroids;
        
        // On doit suivre le compte des vertex dans le cas où la shape contient plusieurs entités
        int nbVertices = 0;
        
        int nbParts = iShape->nParts;
        
        // Dans le cas iShape NULL du coup on continue la boucle
        if( nbParts == 0 )
        {
            SHPDestroyObject( iShape );
            continue;
        }
        
        // On itère sur le nombre d'entités
        for( int j = 0; j < nbParts; ++j )
        {
            // On crée la poly
            AcDbPolyline* poly = new AcDbPolyline();
            
            // On l'ajoute à la Db
            addToModelSpace( poly );
            
            // Nombre de Sommets de la polyligne
            int nbPartVertices;
            
            if( j < nbParts - 1 )
                nbPartVertices = iShape->panPartStart[j + 1] - iShape->panPartStart[j];
            else
                nbPartVertices = iShape->nVertices - iShape->panPartStart[nbParts - 1];
                
            // Boucle sur les vertex
            for( int k = 0; k < nbPartVertices; k++ )
            {
                // On le rajoute à la poly
                poly->addVertexAt( k, AcGePoint2d( iShape->padfX[nbVertices + k], iShape->padfY[nbVertices + k] ) );
            }
            
            // On actualise le nombre de sommets
            nbVertices += nbPartVertices;
            
            // On set les caractéristiques de la polyligne
            if( es = setCharacteristics( hDBF, poly, i, attributesId ) )
            {
                delete poly;
                return es;
            }
            
            // On lit le nom du bloc à insérer
            if( blockFieldIndex != -1 )
                blockName = strToAcStr( DBFReadStringAttribute( hDBF, i, blockFieldIndex ) );
                
            // On ferme la poly
            poly->setClosed( Adesk::kTrue );
            
            // On récupère le centroïde de la poly
            if( getCentroid( poly, ptInsert ) != Acad::eOk )
            {
                SHPDestroyObject( iShape );
                
                for( int k = 0; k < polylines.size(); k++ )
                    polylines[k]->close();
                    
                continue;
            }
            
            // On calcule la bounding box
            AcDbExtents bB;
            poly->getGeomExtents( bB );
            
            // On regarde si le polygone courant est contenu ou contient un des polygones déjà examinés
            int index, isContained;
            
            if( isContainedBy( boundingBoxes, isContained, index, bB ) )
            {
                if( blockName != _T( "" ) )
                {
                    AcGePoint3d pt;
                    
                    // On conserve le bloc à insérer pour le plus grand polygone
                    if( isContained == 1 )
                        pt = centroids[index];
                    else
                        pt = ptInsert;
                        
                    if( insertBlock( hDBF, iShape, i, pt, blockName, attributesId ) )
                    {
                        SHPDestroyObject( iShape );
                        
                        for( int k = 0; k < polylines.size(); k++ )
                            polylines[k]->close();
                            
                        continue;
                    }
                }
                
                if( drawHatch )
                {
                    vector<AcDbPolyline*> borders;
                    borders.push_back( polylines[index] );
                    borders.push_back( poly );
                    
                    if( es = setHatch2d( hDBF, borders, i, attributesId[HATCH], attributesId ) )
                    {
                        SHPDestroyObject( iShape );
                        
                        for( int k = 0; k < polylines.size(); k++ )
                            polylines[k]->close();
                            
                        continue;
                    }
                }
                
                // On ferme/supprime les objets qui ont été ajoutés des vecteurs de stockage
                polylines[index]->close();
                poly->close();
                polylines.erase( polylines.begin() + index );
                centroids.erase( centroids.begin() + index );
                boundingBoxes.erase( boundingBoxes.begin() + index );
                
                nbPolygons++;
            }
            
            // Sinon, on ajoute les informations aux vecteurs de stockage
            else
            {
                polylines.push_back( poly );
                centroids.push_back( ptInsert );
                boundingBoxes.push_back( bB );
            }
        }
        
        // Il reste à parcourir les entités qui n'ont pas encore été traitées
        for( int j = 0; j < polylines.size(); j++ )
        {
            if( blockName != _T( "" ) )
            {
                if( insertBlock( hDBF, iShape, i, centroids[j], blockName, attributesId ) )
                {
                    SHPDestroyObject( iShape );
                    
                    for( int k = 0; k < polylines.size(); k++ )
                        polylines[k]->close();
                        
                    continue;
                }
            }
            
            if( drawHatch )
            {
                vector<AcDbPolyline*> borders;
                borders.push_back( polylines[j] );
                
                if( es = setHatch2d( hDBF, borders, i, attributesId[HATCH], attributesId ) )
                {
                    SHPDestroyObject( iShape );
                    
                    for( int k = 0; k < polylines.size(); k++ )
                        polylines[k]->close();
                        
                    continue;
                }
            }
            
            polylines[j]->close();
            nbPolygons++;
        }
        
        // On supprime la shape
        SHPDestroyObject( iShape );
        
        // On avance la barre de progression
        prog.moveUp( i );
    }
    
    print( "Polygones insérés : %i", nbPolygons );
    
    // Si le bloc entré par l'utilisateur existe, on change le insname après l'insertion
    if( getBlockDefId( blockName ) != AcDbObjectId::kNull )
        setStringVariable( _T( "INSNAME" ), blockName );
        
    return Acad::eOk;
}

void getAngleAndScale( const DBFHandle& hDBF,
    const int& i,
    int( &attributesId )[NB_ATT],
    double& angle,
    double& scaleX,
    double& scaleY,
    double& scaleZ )
{
    // Récupération de la valeur de l'angle
    int angleFieldIndex = attributesId[ANGLE];
    
    if( angleFieldIndex != -1 )
    {
        // On vérifie que l'attribut n'est pas NULL
        if( !DBFIsAttributeNULL( hDBF, i, angleFieldIndex ) )
        {
            // On récupère l'attribut
            angle = DBFReadDoubleAttribute( hDBF, i, angleFieldIndex );
        }
    }
    
    // Récupération de la valeur de l'échelle uniforme
    int scaleFieldIndex = attributesId[SCALE];
    
    if( scaleFieldIndex != -1 )
    {
        // On vérifie que l'attribut n'est pas NULL
        if( !DBFIsAttributeNULL( hDBF, i, scaleFieldIndex ) )
        {
            // On récupère l'attribut
            scaleX = DBFReadDoubleAttribute( hDBF, i, scaleFieldIndex );
            scaleY = scaleX;
            scaleZ = scaleX;
        }
    }
    
    else
    {
        int xscaleFieldIndex = attributesId[XSCALE];
        
        if( xscaleFieldIndex != -1 )
        {
            // On vérifie que l'attribut n'est pas NULL
            if( !DBFIsAttributeNULL( hDBF, i, xscaleFieldIndex ) )
            {
                // On récupère l'attribut
                scaleX = DBFReadDoubleAttribute( hDBF, i, xscaleFieldIndex );
            }
        }
        
        int yscaleFieldIndex = attributesId[YSCALE];
        
        if( yscaleFieldIndex != -1 )
        {
            // On vérifie que l'attribut n'est pas NULL
            if( !DBFIsAttributeNULL( hDBF, i, yscaleFieldIndex ) )
            {
                // On récupère l'attribut
                scaleY = DBFReadDoubleAttribute( hDBF, i, yscaleFieldIndex );
            }
        }
        
        int zscaleFieldIndex = attributesId[ZSCALE];
        
        if( zscaleFieldIndex != -1 )
        {
            // On vérifie que l'attribut n'est pas NULL
            if( !DBFIsAttributeNULL( hDBF, i, zscaleFieldIndex ) )
            {
                // On récupère l'attribut
                scaleZ = DBFReadDoubleAttribute( hDBF, i, zscaleFieldIndex );
            }
        }
    }
}

void convertShpToDwg( char* filePathShp, AcString dir, AcString name, AcString ext )
{
    /// 1. On ouvre le fichier shp
    
    // On ouvre le fichier
    SHPHandle hSHP = SHPOpen( filePathShp, "rb" );
    
    int pnEntities, pnShapeType;
    double adfMinBound[4], adfMaxBound[4];
    
    // On rcupre les info du fichier
    SHPGetInfo( hSHP, &pnEntities, &pnShapeType, adfMinBound, adfMaxBound );
    
    
    /// 2. On récupère les infos du fichier shx
    
    // Chemin du fichier shx
    string filePathShxStr = acStrToStr( dir ) + "\\" + acStrToStr( name ) + ".shx";
    
    // On vérifie que le fichier shx existe
    if( !isFileExisting( filePathShxStr ) )
    {
        // On ferme le fichier shp
        SHPClose( hSHP );
        
        // On informe l'utilisateur que le fichier shx est manquant
        print( "Impossible de trouver le fichier .shx" );
        
        return;
    }
    
    // Sinon
    char* filePathShx = &filePathShxStr[0];
    
    // Ouvrir le fichier d'attributs
    DBFHandle hDBF = DBFOpen( filePathShx, "rb" );
    
    /* On rcupre le rang de chaque attribut d'intrt
    Un attribut d'intrt est dfini dans le header cmdConvertShp.h
    Si un attribut d'intrt est prsent dans le shapefile, attributesId[ATT] renvoie
    le rang correspondant (>=0), sinon renvoie -1*/
    int attributesId[NB_ATT];
    
    for( int i = 0; i < NB_ATT; i++ )
        attributesId[i] = -1;
        
    getAttFieldIndex( hDBF, attributesId );
    
    // On demande le nom du bloc
    AcString blockName = AcString::kEmpty;
    
    if( askBlockName( hSHP,
            hDBF,
            attributesId[BLOCK],
            blockName ) != Acad::eOk )
        return;
        
    // Si polygone, demander si l'on veut hachurer
    bool drawHatch = false;
    
    if( pnShapeType == SHPT_POLYGON || pnShapeType == SHPT_POLYGONZ )
    {
        ACHAR* input;
        AcString initGet = _T( "Oui Non" );
        AcString prompt = _T( "\nHachurer les polygones ? [Oui/<Non>] " );
        acedInitGet( 0, initGet );
        int res = acedGetFullKword( acStrToAcharPointeur( prompt ), input );
        
        if( res == RTNONE )
            input = _T( "Non" );
        else if( res == RTCAN )
            return;
            
        acutPrintf( input );
        
        if( areEqualCase( input, _T( "Oui" ) ) )
            drawHatch = true;
    }
    
    /// 4. On parcourt les shape et on convertit
    
    importShp( hSHP,
        hDBF,
        pnShapeType,
        pnEntities,
        attributesId,
        blockName,
        drawHatch );
        
    // On ferme le fichier shp
    SHPClose( hSHP );
    
    // Fermer le fichier d'attributs
    DBFClose( hDBF );
}

vector<AcString> readParamFile( const AcString& paramFile )
{
    vector<AcString> field;
    
    //Charcger le fichier de configuration
    Book* book = xlCreateBook();
    
    vector<AcString> fildOption[2];
    
    if( book->load( paramFile ) )
    {
        //Recuperer le premier onglet
        Sheet* sheet = book->getSheet( 0 );
        
        if( sheet )
        {
            for( int row = sheet->firstRow(); row < sheet->lastRow(); ++row )
            {
                for( int col = sheet->firstCol(); col < sheet->lastCol(); col++ )
                    fildOption[col].push_back( sheet->readStr( row, col ) );
                    
            }
        }
    }
    
    //On verifie si le troisieme colone de excel n'est pas vide
    vector<AcString>::iterator it = fildOption[1].begin();
    
    //Onboucle sur touts les colones et lines et on prend la bonne valeur de champ
    for( it; it != fildOption[1].end(); ++it )
    {
        //Recuper le numero de la ligne
        int indexLine = it - fildOption[1].begin();
        AcString val = *it;
        
        if( val.compare( "" ) != 0 )
            field.push_back( fildOption[1].at( indexLine ) );
        else
            field.push_back( fildOption[0].at( indexLine ) );
    }
    
    //Fermer le fichier excel
    book->release();
    
    return field;
}

int createFields( DBFHandle& dbfHandle, vector <AcString> fields )
{
    //Boucler sur le string
    vector<AcString>::iterator it = fields.begin();
    
    while( it != fields.end() )
    {
        //Ecrire le champ
        DBFAddField( dbfHandle, *it, FTString, MAXCHAR, 0 );
        ++it;
    }
    
    return fields.end() - fields.begin();
}

void writeEntHandle( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldName )
{
    //Creer un AcDbHandle
    AcString ent_handle = ent->handle().ascii();
    
    //Ecrire le handle de l'entité
    if( DBFGetFieldIndex( dbfHandle, fieldName ) != -1 )
        DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), ent_handle );
}

void writeStatus( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldName )
{
    string stat;
    
    if( ent->isKindOf( AcDbPolyline::desc() ) )
    {
        //Caster l'entité
        AcDbPolyline* poly = AcDbPolyline::cast( ent );
        
        if( isClosed( poly ) )
            stat = "Oui";
        else
            stat = "Non";
    }
    
    else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
    {
        //Caster l'entité
        AcDb3dPolyline* poly = AcDb3dPolyline::cast( ent );
        
        if( isClosed( poly ) )
            stat = "Oui";
        else
            stat = "Non";
    }
    
    else if( ent->isKindOf( AcDbLine::desc() ) )
        stat = "N/A";
    else if( ent->isKindOf( AcDbEllipse::desc() ) )
        stat = "N/A";
    else if( ent->isKindOf( AcDbCircle::desc() ) )
        stat = "N/A";
        
    //Changer en acstring
    auto status = strToAcStr( stat );
    
    //Ecrire le handle de l'entité
    if( DBFGetFieldIndex( dbfHandle, fieldName ) != -1 )
        DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), status );
}

void writeLayerAttribute( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldName )
{
    AcString layer = ent->layer();
    
    //Ecrir le layer
    if( DBFGetFieldIndex( dbfHandle, fieldName ) != -1 )
        DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), layer );
        
}

void writeMaterialName( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldName )
{
    AcString material = ent->material();
    
    //Ecrire le materialname
    if( DBFGetFieldIndex( dbfHandle, fieldName ) != -1 )
        DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), material );
        
}

void writeLength2d( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldName )
{

    //Si : line
    if( ent->isKindOf( AcDbLine::desc() ) )
    {
        //Caster l'entité
        AcDbLine* line = AcDbLine::cast( ent );
        
        //Recuperer la longueur de la ligne
        auto longueur = strToAcStr( to_string( getDistance2d( AcGePoint2d( line->startPoint().x, line->startPoint().y ), AcGePoint2d( line->endPoint().x, line->endPoint().y ) ) ) );
        
        //Ecrire le materialname
        if( DBFGetFieldIndex( dbfHandle, fieldName ) != -1 )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), longueur );
    }
    
    else if( ent->isKindOf( AcDbPolyline::desc() ) )
    {
        //Caster l'entité
        AcDbPolyline* pl = AcDbPolyline::cast( ent );
        
        auto leng = to_string( getLength( pl ) );
        
        //Recuperer la longueur
        auto longueur = strToAcStr( leng );
        
        //Ecrire le materialname
        if( DBFGetFieldIndex( dbfHandle, fieldName ) != -1 )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), longueur );
    }
    
    else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
    {
        //Caster l'entité
        AcDb3dPolyline* pl = AcDb3dPolyline::cast( ent );
        
        auto leng = to_string( get2DLengthPoly( pl ) );
        
        //Recuperer la longueur
        auto longueur = strToAcStr( leng );
        
        //Ecrire le materialname
        if( DBFGetFieldIndex( dbfHandle, fieldName ) != -1 )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), longueur );
    }
    
    else if( ent->isKindOf( AcDbCircle::desc() ) )
    {
    
        //Recuperer la longueur
        auto longueur = "N/A";
        
        //Ecrire le materialname
        if( DBFGetFieldIndex( dbfHandle, fieldName ) != -1 )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), longueur );
    }
    
}

void writeLength3d( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldName )
{

    if( ent->isKindOf( AcDb3dPolyline::desc() ) )
    {
        //Caster l'entité
        AcDb3dPolyline* pl = AcDb3dPolyline::cast( ent );
        
        //Recuperer la longueur
        auto longueur = strToAcStr( to_string( getLengthPoly( pl ) ) );
        
        //Ecrire le materialname
        if( DBFGetFieldIndex( dbfHandle, fieldName ) != -1 )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), longueur );
    }
    
    else  if( ent->isKindOf( AcDbPolyline::desc() ) )
    {
        //Caster l'entité
        AcDbPolyline* pl = AcDbPolyline::cast( ent );
        
        //Recuperer la longueur
        auto longueur = strToAcStr( to_string( getLength( pl ) ) );
        
        //Ecrire le materialname
        if( DBFGetFieldIndex( dbfHandle, fieldName ) != -1 )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), longueur );
    }
    
    else  if( ent->isKindOf( AcDbCircle::desc() ) )
    {
        //Recuperer la longueur
        auto longueur = "N/A";
        
        //Ecrire le materialname
        if( DBFGetFieldIndex( dbfHandle, fieldName ) != -1 )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), longueur );
    }
    
    else  if( ent->isKindOf( AcDbLine::desc() ) )
    {
        //Caster l'entité
        AcDbLine* line = AcDbLine::cast( ent );
        
        AcGePoint2d pstart = AcGePoint2d( line->startPoint().x, line->startPoint().y );
        AcGePoint2d pend = AcGePoint2d( line->endPoint().x, line->endPoint().y )  ;
        
        //Recuperer la longueur de la ligne
        auto longueur = strToAcStr( to_string( getDistance2d( pstart, pend ) ) ) ;
        
        //Ecrire le materialname
        if( DBFGetFieldIndex( dbfHandle, fieldName ) != -1 )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), longueur );
    }
    
    
}




void writeColorAttribute( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldName )
{
    AcCmColor acColor = ent->color();
    
    //Recuper la couleur en RGB
    int red = acColor.red();
    int green = acColor.green();
    int blue = acColor.blue();
    
    AcString value = numberToAcString( red ) +
        "," +
        numberToAcString( green ) +
        "," +
        numberToAcString( blue );
        
    if( acColor.isByLayer() )
        value = "ByLayer";
        
    if( acColor.isByPen() )
        value = "ByPen";
        
    if( acColor.isByBlock() )
        value = "ByBlock";
        
    DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), value );
    
}
void writePolyHandle( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldName )
{
    if( ent->isKindOf( AcDbPolyline::desc() ) )
    {
        //Caster l'entité;
        AcDbPolyline* poly = AcDbPolyline::cast( ent );
        
        //Recuperer le handle de la polyline
        auto poly_handle = poly->handle().ascii();
        
        //Inserer dans le dbfile
        
        if( DBFGetFieldIndex( dbfHandle, fieldName ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), poly_handle );
            
    }
    
    else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
    {
        //Caster l'entité;
        AcDb3dPolyline* poly = AcDb3dPolyline::cast( ent );
        
        //Recuperer le handle de la polyline
        auto poly_handle = poly->handle().ascii();
        
        //Inserer dans le dbfile
        
        if( DBFGetFieldIndex( dbfHandle, fieldName ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), poly_handle );
            
    }
}

void writeLynetypeAttribute( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldName )
{
    AcString lineType = ent->linetype();
    
    DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), lineType );
}

void writeLineTypeScaleAttribute( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldName )
{
    double lineScale = ent->linetypeScale();
    
    DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), strToAcStr( to_string( lineScale ) ) );
}

void writeLineWeightAttribute( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldName )
{
    //
    AcDb::LineWeight lineWeight = ent->lineWeight();
    
    DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), numberToAcString( lineWeight ) );
}

void writeTransparency( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldName )
{
    AcCmTransparency transparency = ent->transparency();
    
    if( transparency.isByAlpha() ||
        transparency.isSolid() ||
        transparency.isClear() )
    {
        double transparence = transparency.alpha();
        DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), strToAcStr( to_string( transparence ) ) );
        
    }
    
    else if( transparency.isByBlock() )
    {
        AcString transparence = _T( "ByBlock" );
        DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), transparence );
    }
    
    else if( transparency.isByLayer() )
    {
        AcString transparence = _T( "ByLayer" );
        DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), transparence );
    }
}

void writeBlockName( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldName )
{
    //Tester si l'entite est un blokc
    if( ent->isKindOf( AcDbBlockReference::desc() ) )
    {
        // On cast l'entité du fichier de départ en bloc
        AcDbBlockReference* pBlock = AcDbBlockReference::cast( ent );
        
        AcString blockname;
        Acad::ErrorStatus es = getBlockName( blockname, pBlock );
        
        if( es )
            print( es );
        else
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), blockname );
    }
}

void writeOrientation( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldName )
{
    //Tester si l'entite est un blokc
    if( ent->isKindOf( AcDbBlockReference::desc() ) )
    {
    
        AcDbBlockReference* pBlock = AcDbBlockReference::cast( ent );
        
        double angle = pBlock->rotation();
        
        DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), strToAcStr( to_string( angle ) ) );
    }
    
    else if( ent->isKindOf( AcDbText::desc() ) )
    {
        AcDbText* txt = AcDbText::cast( ent );
        
        double angle = txt->rotation();
        DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), strToAcStr( to_string( angle ) ) );
        
    }
}

void writeXPos( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldname )
{
    //Position x
    AcString pos;
    
    //Si l'entité est un texte
    if( ent->isKindOf( AcDbText::desc() ) )
    {
        //Caster l'entité
        AcDbText* txt = AcDbText::cast( ent );
        
        //Determiner la position
        pos = strToAcStr( to_string( txt->position().x ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), pos );
    }
    
    //Si l'entité est un mtext
    if( ent->isKindOf( AcDbMText::desc() ) )
    {
        //Caster l'entité
        AcDbMText* mtxt = AcDbMText::cast( ent );
        
        //Determiner la position
        pos = strToAcStr( to_string( mtxt->location().x ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), pos );
    }
    
    else if( ent->isKindOf( AcDbBlockReference::desc() ) )
    {
        //caster l'entité
        AcDbBlockReference* blk = AcDbBlockReference::cast( ent );
        
        //Determiner la position
        pos = strToAcStr( to_string( blk->position().x ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), pos );
    }
    
    else if( ent->isKindOf( AcDbCircle::desc() ) )
    {
        //Caster l'entité
        AcDbCircle* circle = AcDbCircle::cast( ent );
        
        //Determiner la position
        pos = strToAcStr( to_string( circle->center().x ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), pos );
    }
    
    else if( ent->isKindOf( AcDbHatch::desc() ) )
    {
        //Caster l'entité
        AcDbHatch* hatch = AcDbHatch::cast( ent );
        
        //Recuperer son boundingbox
        AcDbExtents ext;
        hatch->getGeomExtents( ext );
        
        //Determiner la position du centre
        AcGePoint3d ptCenter = midPoint3d( ext.minPoint(), ext.maxPoint() );
        
        //Recuperer la position
        pos = strToAcStr( to_string( ptCenter.x ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), pos );
    }
    
    else if( ent->isKindOf( AcDbPoint::desc() ) )
    {
        //Caster l'entité
        AcDbPoint* pt = AcDbPoint::cast( ent );
        
        //Determiner la position
        pos = strToAcStr( to_string( pt->position().x ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), pos );
    }
}

void writeXScale( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldName )
{
    if( ent->isKindOf( AcDbBlockReference::desc() ) )
    {
        AcDbBlockReference* pBlock = AcDbBlockReference::cast( ent );
        
        double xScale = pBlock->scaleFactors().sx;
        
        DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), strToAcStr( to_string( xScale ) ) );
        
    }
}

void writeYPos( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldname )
{
    //Position y
    AcString pos;
    
    //Si l'entité est un texte
    if( ent->isKindOf( AcDbText::desc() ) )
    {
        //Caster l'entité
        AcDbText* txt = AcDbText::cast( ent );
        
        //Determiner la position
        pos = strToAcStr( to_string( txt->position().y ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), pos );
    }
    
    //Si l'entité est un mtext
    if( ent->isKindOf( AcDbMText::desc() ) )
    {
        //Caster l'entité
        AcDbMText* mtxt = AcDbMText::cast( ent );
        
        //Determiner la position
        pos = strToAcStr( to_string( mtxt->location().y ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), pos );
    }
    
    else if( ent->isKindOf( AcDbBlockReference::desc() ) )
    {
        //caster l'entité
        AcDbBlockReference* blk = AcDbBlockReference::cast( ent );
        
        //Determiner la position
        pos = strToAcStr( to_string( blk->position().y ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), pos );
    }
    
    else if( ent->isKindOf( AcDbCircle::desc() ) )
    {
        //Caster l'entité
        AcDbCircle* circle = AcDbCircle::cast( ent );
        
        //Determiner la position
        pos = strToAcStr( to_string( circle->center().y ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), pos );
    }
    
    else if( ent->isKindOf( AcDbHatch::desc() ) )
    {
        //Caster l'entité
        AcDbHatch* hatch = AcDbHatch::cast( ent );
        
        //Recuperer son boundingbox
        AcDbExtents ext;
        hatch->getGeomExtents( ext );
        
        //Determiner la position du centre
        AcGePoint3d ptCenter = midPoint3d( ext.minPoint(), ext.maxPoint() );
        
        //Recuperer la position
        pos = strToAcStr( to_string( ptCenter.y ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), pos );
    }
    
    else if( ent->isKindOf( AcDbPoint::desc() ) )
    {
        //Caster l'entité
        AcDbPoint* pt = AcDbPoint::cast( ent );
        
        //Determiner la position
        pos = strToAcStr( to_string( pt->position().y ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), pos );
    }
}

void writeYScale( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldName )
{
    if( ent->isKindOf( AcDbBlockReference::desc() ) )
    {
        AcDbBlockReference* pBlock = AcDbBlockReference::cast( ent );
        
        double yScale = pBlock->scaleFactors().sy;
        
        DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), strToAcStr( to_string( yScale ) ) );
    }
}

void writeZPos( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldname )
{
    //Position z
    AcString pos;
    
    //Si l'entité est un texte
    if( ent->isKindOf( AcDbText::desc() ) )
    {
        //Caster l'entité
        AcDbText* txt = AcDbText::cast( ent );
        
        //Determiner la position
        pos = strToAcStr( to_string( txt->position().z ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), pos );
    }
    
    //Si l'entité est un mtext
    if( ent->isKindOf( AcDbMText::desc() ) )
    {
        //Caster l'entité
        AcDbMText* mtxt = AcDbMText::cast( ent );
        
        //Determiner la position
        pos = strToAcStr( to_string( mtxt->location().z ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), pos );
    }
    
    else if( ent->isKindOf( AcDbBlockReference::desc() ) )
    {
        //caster l'entité
        AcDbBlockReference* blk = AcDbBlockReference::cast( ent );
        
        //Determiner la position
        pos = strToAcStr( to_string( blk->position().z ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), pos );
    }
    
    else if( ent->isKindOf( AcDbCircle::desc() ) )
    {
        //Caster l'entité
        AcDbCircle* circle = AcDbCircle::cast( ent );
        
        //Determiner la position
        pos = strToAcStr( to_string( circle->center().z ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), pos );
    }
    
    else if( ent->isKindOf( AcDbHatch::desc() ) )
    {
        //Caster l'entité
        AcDbHatch* hatch = AcDbHatch::cast( ent );
        
        //Recuperer son boundingbox
        AcDbExtents ext;
        hatch->getGeomExtents( ext );
        
        //Determiner la position du centre
        AcGePoint3d ptCenter = midPoint3d( ext.minPoint(), ext.maxPoint() );
        
        //Recuperer la position
        pos = strToAcStr( to_string( ptCenter.z ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), pos );
    }
    
    else if( ent->isKindOf( AcDbPoint::desc() ) )
    {
        //Caster l'entité
        AcDbPoint* pt = AcDbPoint::cast( ent );
        
        //Determiner la position
        pos = strToAcStr( to_string( pt->position().z ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), pos );
    }
}

void writeArea( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldname )
{
    //Position z
    AcString area_string;
    
    if( ent->isKindOf( AcDbCircle::desc() ) )
    {
        //Caster l'entité
        AcDbCircle* circle = AcDbCircle::cast( ent );
        
        double area;
        circle->getArea( area ) ;
        
        //Determiner la position
        area_string = strToAcStr( to_string( area ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), area_string );
    }
    
    else if( ent->isKindOf( AcDbHatch::desc() ) )
    {
        //Caster l'entité
        AcDbHatch* hatch = AcDbHatch::cast( ent );
        
        double area;
        hatch->getArea( area );
        
        //Determiner la position
        area_string = strToAcStr( to_string( area ) );
        
        //Ecrire le resultat dans le dbfile
        DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), area_string );
    }
    
    else if( ent->isKindOf( AcDbPolyline::desc() ) )
    {
        //Caster l'entité
        AcDbPolyline* pt = AcDbPolyline::cast( ent );
        
        double area;
        
        pt->getArea( area );
        
        //Determiner la surface
        area_string = strToAcStr( to_string( area ) );
        
        //Ecrire le resultat dans le dbfile
        DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), area_string );
    }
    
    else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
    {
        //Caster l'entité
        AcDb3dPolyline* pt = AcDb3dPolyline::cast( ent );
        
        double area;
        
        //Recuperer la surface
        pt->getArea( area );
        
        //Determiner la surface
        area_string = strToAcStr( to_string( area ) );
        
        //Ecrire le resultat dans le dbfile
        DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), area_string );
    }
}

void writePattern( AcDbHatch* ent, DBFHandle dbfHandle, int iShape, AcString fieldname )
{
    //Définir le motif
    AcString motif = ent->patternName();
    
    //On verifie que la colone pour patern existe
    if( DBFGetFieldIndex( dbfHandle, fieldname ) )
        DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), motif );
        
}

void writeThicknessAttribute( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldname )
{
    //Position z
    double thickness;
    
    //Si l'entité est un texte
    if( ent->isKindOf( AcDbText::desc() ) )
    {
        //Caster l'entité
        AcDbText* txt = AcDbText::cast( ent );
        
        //Determiner la position
        thickness =  txt->thickness();
        auto thick = strToAcStr( to_string( thickness ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), thick );
    }
    
    
    else if( ent->isKindOf( AcDbCircle::desc() ) )
    {
        //Caster l'entité
        AcDbCircle* circle = AcDbCircle::cast( ent );
        auto thick = strToAcStr( to_string( thickness ) );
        
        //Determiner la position
        thickness = circle->thickness();
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), thick );
    }
    
    else if( ent->isKindOf( AcDbPoint::desc() ) )
    {
        //Caster l'entité
        AcDbPoint* pt = AcDbPoint::cast( ent );
        
        //Determiner la position
        thickness = pt->thickness();
        auto thick = strToAcStr( to_string( thickness ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), thick );
    }
    
    else if( ent->isKindOf( AcDbPolyline::desc() ) )
    {
        //Caster l'entité
        AcDbPolyline* pt = AcDbPolyline::cast( ent );
        
        //Determiner la position
        thickness = pt->thickness();
        auto thick = strToAcStr( to_string( thickness ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), thick );
    }
    
    else if( ent->isKindOf( AcDbLine::desc() ) )
    {
        //Caster l'entité
        AcDbLine* pt = AcDbLine::cast( ent );
        
        //Determiner la position
        thickness = pt->thickness();
        auto thick = strToAcStr( to_string( thickness ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), thick );
    }
    
    else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
    {
        auto thick = strToAcStr( "N/A" );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), thick );
    }
    
    else if( ent->isKindOf( AcDbEllipse::desc() ) )
    {
    
        auto thick = strToAcStr( "N/A" );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), thick );
    }
    
    
}

void writeLineStartX( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldname )
{

    if( ent->isKindOf( AcDbLine::desc() ) )
    {
        //Caster l'entité
        AcDbLine* line = AcDbLine::cast( ent );
        
        //Recuperer le point de depart de la line
        AcGePoint3d startPt = line->startPoint();
        
        //Recuperer x du startpoint
        double temp = startPt.x;
        
        //changer en acstring
        auto start = strToAcStr( to_string( temp ) ) ;
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), start );
    }
    
}

void writeLineStartY( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldname )
{
    if( ent->isKindOf( AcDbLine::desc() ) )
    {
        //Caster l'entité
        AcDbLine* line = AcDbLine::cast( ent );
        
        //Recuperer y du startpoint
        double temp = line->startPoint().y;
        
        //changer en acstring
        auto start = strToAcStr( to_string( temp ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), start );
    }
}

void writeLineStartZ( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldname )
{
    if( ent->isKindOf( AcDbLine::desc() ) )
    {
        //Caster l'entité
        AcDbLine* line = AcDbLine::cast( ent );
        
        //Recuperer z du startpoint
        double temp = line->startPoint().z;
        
        //changer en acstring
        auto start = strToAcStr( to_string( temp ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), start );
    }
    
}

void writeLineEndX( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldname )
{
    if( ent->isKindOf( AcDbLine::desc() ) )
    {
        //Caster l'entité
        AcDbLine* line = AcDbLine::cast( ent );
        
        //Recuperer x du endPoint
        double temp = line->endPoint().x;
        
        //changer en acstring
        auto end = strToAcStr( to_string( temp ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), end );
    }
    
}

void writeLineEndY( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldname )
{
    if( ent->isKindOf( AcDbLine::desc() ) )
    {
        //Caster l'entité
        AcDbLine* line = AcDbLine::cast( ent );
        
        ///Recuperer x du endPoint
        double temp = line->endPoint().y;
        
        //changer en acstring
        auto end = strToAcStr( to_string( temp ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), end );
    }
}

void writeLineEndZ( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldname )
{
    if( ent->isKindOf( AcDbLine::desc() ) )
    {
        //Caster l'entité
        AcDbLine* line = AcDbLine::cast( ent );
        
        //Recuperer x du endPoint
        double temp = line->endPoint().z;
        
        //changer en acstring
        auto end = strToAcStr( to_string( temp ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldname ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldname ), end );
    }
}

void writeRotation( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldName )
{
    if( ent->isKindOf( AcDbBlockReference::desc() ) )
    {
        //Caster l'entité
        AcDbBlockReference* blk = AcDbBlockReference::cast( ent );
        
        //Determiner la rotation du block
        AcString rotation = strToAcStr( to_string( blk->rotation() ) );
        
        //Ecrire le resultat dans le dbfile
        if( DBFGetFieldIndex( dbfHandle, fieldName ) )
            DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), rotation );
            
    }
}

void writeZScale( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, AcString fieldName )
{
    if( ent->isKindOf( AcDbBlockReference::desc() ) )
    {
        AcDbBlockReference* pBlock = AcDbBlockReference::cast( ent );
        
        double zScale = pBlock->scaleFactors().sz;
        
        DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, fieldName ), strToAcStr( to_string( zScale ) ) );
    }
}

void writeBlock( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, vector<AcString> fieldName )
{
    writeEntHandle( ent, dbfHandle, iShape, fieldName.at( 0 ) );
    writeLayerAttribute( ent, dbfHandle, iShape, fieldName.at( 1 ) );
    writeColorAttribute( ent, dbfHandle, iShape, fieldName.at( 2 ) );
    writeLynetypeAttribute( ent, dbfHandle, iShape, fieldName.at( 3 ) );
    writeLineTypeScaleAttribute( ent, dbfHandle, iShape, fieldName.at( 4 ) );
    writeLineWeightAttribute( ent, dbfHandle, iShape, fieldName.at( 5 ) );
    writeTransparency( ent, dbfHandle, iShape, fieldName.at( 6 ) );
    writeBlockName( ent, dbfHandle, iShape, fieldName.at( find( fieldName.begin(), fieldName.end(), _T( "BLOCK" ) ) - fieldName.begin() ) );
    writeMaterialName( ent, dbfHandle, iShape, fieldName.at( 8 ) );
    writeXPos( ent, dbfHandle, iShape, fieldName.at( 9 ) );
    writeYPos( ent, dbfHandle, iShape, fieldName.at( 10 ) );
    writeZPos( ent, dbfHandle, iShape, fieldName.at( 11 ) );
    writeRotation( ent, dbfHandle, iShape, fieldName.at( 12 ) );
    writeAttribute( ent, dbfHandle, iShape, fieldName );
    
}

void writeBlockPoly( AcDbEntity* ent, AcDbEntity* entityPoly, DBFHandle dbfHandle, int iShape, vector<AcString> fieldName )
{
    writeEntHandle( ent, dbfHandle, iShape, fieldName.at( 0 ) );
    writeLayerAttribute( ent, dbfHandle, iShape, fieldName.at( 1 ) );
    writeColorAttribute( ent, dbfHandle, iShape, fieldName.at( 2 ) );
    writeLynetypeAttribute( ent, dbfHandle, iShape, fieldName.at( 3 ) );
    writeLineTypeScaleAttribute( ent, dbfHandle, iShape, fieldName.at( 4 ) );
    writeLineWeightAttribute( ent, dbfHandle, iShape, fieldName.at( 5 ) );
    writeTransparency( ent, dbfHandle, iShape, fieldName.at( 6 ) );
    writeBlockName( ent, dbfHandle, iShape, fieldName.at( find( fieldName.begin(), fieldName.end(), _T( "BLOCK" ) ) - fieldName.begin() ) );
    writeMaterialName( ent, dbfHandle, iShape, fieldName.at( 8 ) );
    writeXPos( ent, dbfHandle, iShape, fieldName.at( 9 ) );
    writeYPos( ent, dbfHandle, iShape, fieldName.at( 10 ) );
    writeZPos( ent, dbfHandle, iShape, fieldName.at( 11 ) );
    writeRotation( ent, dbfHandle, iShape, fieldName.at( 12 ) );
    writeEntHandle( entityPoly, dbfHandle, iShape, fieldName.at( 13 ) );
    writeLength2d( entityPoly, dbfHandle, iShape, fieldName.at( 14 ) );
    writeLength3d( entityPoly, dbfHandle, iShape, fieldName.at( 15 ) );
    writeArea( entityPoly, dbfHandle, iShape, fieldName.at( 16 ) );
    writeAttribute( ent, dbfHandle, iShape, fieldName );
    
}

void writeAttribute( AcDbEntity* ent, DBFHandle dbfHandle, int iShape, vector<AcString> fieldVect )
{
    if( ent->isKindOf( AcDbBlockReference::desc() ) )
    {
        //Recuberer le block reference
        AcDbBlockReference* pBlock = AcDbBlockReference::cast( ent );
        map<AcString, AcString> att = getBlockAttWithValuesList( pBlock );
        map<AcString, AcString>::iterator it = att.begin();
        
        //Boucler et ecrire sur les entite
        while( it != att.end() )
        {
            string tempField = acStrToStr( it->first );
            AcString value = it->second;
            
            for( int i = 10; i <= tempField.length(); i++ )
                tempField.pop_back();
                
            AcString field = strToAcStr( tempField );
            
            if( value.isEmpty() )
                value = "";
                
            if( field.isEmpty() )
                return;
                
            if( DBFGetFieldIndex( dbfHandle, field ) != -1 )
                DBFWriteStringAttribute( dbfHandle, iShape, DBFGetFieldIndex( dbfHandle, field ), value );
                
            it++;
        }
    }
}

int createObjectLine( AcDbEntity* ent, double*& x, double*& y, double* z )
{
    //Verifier si c'est un ligne
    if( ent->isKindOf( AcDbLine::desc() ) )
    {
        //Reciperer l'objet ligne
        AcDbLine* line = AcDbLine::cast( ent );
        
        //Coordonne du polygone
        *x = line->startPoint().x;
        *y = line->startPoint().y;
        *z = line->startPoint().z;
        
        *( x + 1 ) = line->startPoint().x;
        *( y + 1 ) = line->startPoint().y;
        *( z + 1 ) = line->startPoint().z;
    }
    
    //Si l'entite est un polyligne2d
    else if( ent->isKindOf( AcDbPolyline::desc() ) )
    {
        AcDbPolyline* poly2d = AcDbPolyline::cast( ent );
        
        //Boucler sur les sommets
        for( int j = 0; j < poly2d->numVerts(); j++ )
        {
            AcGePoint3d pt;
            poly2d->getPointAt( j, pt );
            *( x + j ) = pt.x;
            *( y + j ) = pt.y;
            *( z + j ) = pt.z;
        }
    }
    
    //Si c'est un polyligne 3d
    else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
    {
        AcDb3dPolyline* poly3d = AcDb3dPolyline::cast( ent );
        
        //Boucler sur les sommets
        AcDbObjectIterator* iterPoly = poly3d->vertexIterator();
        AcDb3dPolylineVertex* vertex;
        int j = 0;
        
        for( iterPoly->start(); !iterPoly->done(); iterPoly->step() )
        {
            if( Acad::eOk == poly3d->openVertex( vertex, iterPoly->objectId(), AcDb::kForRead ) )
            {
                *( x + j ) = vertex->position().x;
                *( y + j ) = vertex->position().y;
                *( z + j ) = vertex->position().z;
            }
            
            j++;
        }
    }
    
    return sizeof( y );
}

void writeConsantWidht( AcDbEntity* ent, DBFHandle dBFHandle, int iShape, AcString field )
{
    if( ent->isKindOf( AcDbPolyline::desc() ) )
    {
        AcDbPolyline* poly = AcDbPolyline::cast( ent );
        
        double cWidth = poly->getConstantWidth();
        
        if( DBFGetFieldIndex( dBFHandle, field ) )
            DBFWriteDoubleAttribute( dBFHandle, iShape, DBFGetFieldIndex( dBFHandle, field ), cWidth );
    }
    
}

void writeLine( AcDbEntity* ent, DBFHandle dBFHandle, const int& iShape, const vector<AcString>& fieldName )
{
    writeEntHandle( ent, dBFHandle, iShape, fieldName.at( 0 ) );
    writeLayerAttribute( ent, dBFHandle, iShape, fieldName.at( 1 ) );
    writeColorAttribute( ent, dBFHandle, iShape, fieldName.at( 2 ) );
    writeLynetypeAttribute( ent, dBFHandle, iShape, fieldName.at( 3 ) );
    writeLineTypeScaleAttribute( ent, dBFHandle, iShape, fieldName.at( 4 ) );
    writeLineWeightAttribute( ent, dBFHandle, iShape, fieldName.at( 5 ) );
    writeTransparency( ent, dBFHandle, iShape, fieldName.at( 6 ) );
    writeThicknessAttribute( ent, dBFHandle, iShape, fieldName.at( 7 ) );
    writeMaterialName( ent, dBFHandle, iShape, fieldName.at( 8 ) );
    writeLength2d( ent, dBFHandle, iShape, fieldName.at( 9 ) );
    writeLineStartX( ent, dBFHandle, iShape, fieldName.at( 10 ) );
    writeLineStartY( ent, dBFHandle, iShape, fieldName.at( 11 ) );
    writeLineStartZ( ent, dBFHandle, iShape, fieldName.at( 12 ) );
    writeLineEndX( ent, dBFHandle, iShape, fieldName.at( 13 ) );
    writeLineEndY( ent, dBFHandle, iShape, fieldName.at( 14 ) );
    writeLineEndZ( ent, dBFHandle, iShape, fieldName.at( 15 ) );
}

void writeCurve( AcDbEntity* ent, DBFHandle dBFHandle, const int& iShape, const vector<AcString>& fieldName )
{
    writeEntHandle( ent, dBFHandle, iShape, fieldName.at( 0 ) );
    writeLayerAttribute( ent, dBFHandle, iShape, fieldName.at( 1 ) );
    writeColorAttribute( ent, dBFHandle, iShape, fieldName.at( 2 ) );
    writeLynetypeAttribute( ent, dBFHandle, iShape, fieldName.at( 3 ) );
    writeLineTypeScaleAttribute( ent, dBFHandle, iShape, fieldName.at( 4 ) );
    writeLineWeightAttribute( ent, dBFHandle, iShape, fieldName.at( 5 ) );
    writeTransparency( ent, dBFHandle, iShape, fieldName.at( 6 ) );
    writeThicknessAttribute( ent, dBFHandle, iShape, fieldName.at( 7 ) );
    writeMaterialName( ent, dBFHandle, iShape, fieldName.at( 8 ) );
    writeLength2d( ent, dBFHandle, iShape, fieldName.at( 9 ) );
    writeLength3d( ent, dBFHandle, iShape, fieldName.at( 10 ) );
    writeStatus( ent, dBFHandle, iShape, fieldName.at( 11 ) );
    
}

void writePoly2d( AcDbEntity* ent, DBFHandle dBFHandle, const int& iShape, const vector<AcString>& fieldName )
{
    writeEntHandle( ent, dBFHandle, iShape, fieldName.at( 0 ) );
    writeLayerAttribute( ent, dBFHandle, iShape, fieldName.at( 1 ) );
    writeColorAttribute( ent, dBFHandle, iShape, fieldName.at( 2 ) );
    writeLynetypeAttribute( ent, dBFHandle, iShape, fieldName.at( 3 ) );
    writeLineTypeScaleAttribute( ent, dBFHandle, iShape, fieldName.at( 4 ) );
    writeLineWeightAttribute( ent, dBFHandle, iShape, fieldName.at( 5 ) );
    writeTransparency( ent, dBFHandle, iShape, fieldName.at( 6 ) );
    writeThicknessAttribute( ent, dBFHandle, iShape, fieldName.at( 7 ) );
    writeMaterialName( ent, dBFHandle, iShape, fieldName.at( 8 ) );
    writeLength2d( ent, dBFHandle, iShape, fieldName.at( 9 ) );
    writeLength3d( ent, dBFHandle, iShape, fieldName.at( 10 ) );
    writeStatus( ent, dBFHandle, iShape, fieldName.at( 11 ) );
}

void writeCircle( AcDbEntity* ent, DBFHandle dBFHandle, const int& iShape, const vector<AcString>& fieldName )
{
    writeEntHandle( ent, dBFHandle, iShape, fieldName.at( 0 ) );
    writeLayerAttribute( ent, dBFHandle, iShape, fieldName.at( 1 ) );
    writeColorAttribute( ent, dBFHandle, iShape, fieldName.at( 2 ) );
    writeLynetypeAttribute( ent, dBFHandle, iShape, fieldName.at( 3 ) );
    writeLineTypeScaleAttribute( ent, dBFHandle, iShape, fieldName.at( 4 ) );
    writeLineWeightAttribute( ent, dBFHandle, iShape, fieldName.at( 5 ) );
    writeTransparency( ent, dBFHandle, iShape, fieldName.at( 6 ) );
    writeThicknessAttribute( ent, dBFHandle, iShape, fieldName.at( 7 ) );
    writeMaterialName( ent, dBFHandle, iShape, fieldName.at( 8 ) );
    writeXPos( ent, dBFHandle, iShape, fieldName.at( 9 ) );
    writeYPos( ent, dBFHandle, iShape, fieldName.at( 10 ) );
    writeZPos( ent, dBFHandle, iShape, fieldName.at( 11 ) );
    writeRadius( ent, dBFHandle, iShape, fieldName.at( 12 ) );
    
}


void writeClosedPoly2d( AcDbEntity* ent, DBFHandle dBFHandle, const int& iShape, const vector<AcString>& fieldName )
{
    writeEntHandle( ent, dBFHandle, iShape, fieldName.at( 0 ) );
    writeLayerAttribute( ent, dBFHandle, iShape, fieldName.at( 1 ) );
    writeColorAttribute( ent, dBFHandle, iShape, fieldName.at( 2 ) );
    writeLynetypeAttribute( ent, dBFHandle, iShape, fieldName.at( 3 ) );
    writeLineTypeScaleAttribute( ent, dBFHandle, iShape, fieldName.at( 4 ) );
    writeLineWeightAttribute( ent, dBFHandle, iShape, fieldName.at( 5 ) );
    writeTransparency( ent, dBFHandle, iShape, fieldName.at( 6 ) );
    writeThicknessAttribute( ent, dBFHandle, iShape, fieldName.at( 7 ) );
    writeMaterialName( ent, dBFHandle, iShape, fieldName.at( 8 ) );
    writeLength2d( ent, dBFHandle, iShape, fieldName.at( 9 ) );
    writeLength2d( ent, dBFHandle, iShape, fieldName.at( 10 ) );
    writeArea( ent, dBFHandle, iShape, fieldName.at( 11 ) );
}

void writePoly3d( AcDbEntity* ent, DBFHandle dBFHandle, const int& iShape, const vector<AcString>& fieldName )
{
    writeEntHandle( ent, dBFHandle, iShape, fieldName.at( 0 ) );
    writeLayerAttribute( ent, dBFHandle, iShape, fieldName.at( 1 ) );
    writeColorAttribute( ent, dBFHandle, iShape, fieldName.at( 2 ) );
    writeLynetypeAttribute( ent, dBFHandle, iShape, fieldName.at( 3 ) );
    writeLineTypeScaleAttribute( ent, dBFHandle, iShape, fieldName.at( 4 ) );
    writeLineWeightAttribute( ent, dBFHandle, iShape, fieldName.at( 5 ) );
    writeTransparency( ent, dBFHandle, iShape, fieldName.at( 6 ) );
    writeThicknessAttribute( ent, dBFHandle, iShape, fieldName.at( 7 ) );
    writeMaterialName( ent, dBFHandle, iShape, fieldName.at( 8 ) );
    writeLength2d( ent, dBFHandle, iShape, fieldName.at( 9 ) );
    writeLength3d( ent, dBFHandle, iShape, fieldName.at( 10 ) );
    writeStatus( ent, dBFHandle, iShape, fieldName.at( 11 ) );
}

void writeClosedPoly3d( AcDbEntity* ent, DBFHandle dBFHandle, const int& iShape, const vector<AcString>& fieldName )
{
    writeEntHandle( ent, dBFHandle, iShape, fieldName.at( 0 ) );
    writeLayerAttribute( ent, dBFHandle, iShape, fieldName.at( 1 ) );
    writeColorAttribute( ent, dBFHandle, iShape, fieldName.at( 2 ) );
    writeLynetypeAttribute( ent, dBFHandle, iShape, fieldName.at( 3 ) );
    writeLineTypeScaleAttribute( ent, dBFHandle, iShape, fieldName.at( 4 ) );
    writeLineWeightAttribute( ent, dBFHandle, iShape, fieldName.at( 5 ) );
    writeTransparency( ent, dBFHandle, iShape, fieldName.at( 6 ) );
    writeThicknessAttribute( ent, dBFHandle, iShape, fieldName.at( 7 ) );
    writeMaterialName( ent, dBFHandle, iShape, fieldName.at( 8 ) );
    writeLength2d( ent, dBFHandle, iShape, fieldName.at( 9 ) );
    writeLength3d( ent, dBFHandle, iShape, fieldName.at( 10 ) );
    writeArea( ent, dBFHandle, iShape, fieldName.at( 11 ) );
}

void writeTextStr( AcDbEntity* ent, DBFHandle dBFHandle, int iShape, AcString field )
{
    if( ent->isKindOf( AcDbText::desc() ) )
    {
    
        AcDbText* text = AcDbText::cast( ent );
        AcString txt = text->textString();
        
        //Verifie si le champ existe
        if( DBFGetFieldIndex( dBFHandle, field ) )
            DBFWriteStringAttribute( dBFHandle, iShape, DBFGetFieldIndex( dBFHandle, field ), txt );
            
    }
}

void writeTexteHeight( AcDbEntity* ent, DBFHandle dBFHandle, int iShape, AcString field )
{
    if( ent->isKindOf( AcDbText::desc() ) )
    {
        AcDbText* text = AcDbText::cast( ent );
        double height = text->height();
        
        //Verifie si le champ existe
        if( DBFGetFieldIndex( dBFHandle, field ) )
            DBFWriteStringAttribute( dBFHandle, iShape, DBFGetFieldIndex( dBFHandle, field ), strToAcStr( to_string( height ) ) );
    }
}

void writeTextAlign( AcDbEntity* ent, DBFHandle dBFHandle, int iShape, AcString field )
{
    if( ent->isKindOf( AcDbText::desc() ) )
    {
        AcDbText* texte = AcDbText::cast( ent );
        AcDb::TextHorzMode hortzMode = texte->horizontalMode();
        AcString hortzM, vertM;
        
        switch( hortzMode )
        {
            case AcDb::kTextLeft:
                hortzM = "A gauche";
                break;
                
            case AcDb::kTextCenter:
                hortzM = "Centrer";
                break;
                
            case AcDb::kTextRight:
                hortzM = "A droite";
                break;
                
            case AcDb::kTextAlign:
                hortzM = "Aligne";
                break;
                
            case AcDb::kTextMid:
                hortzM = "Milieu";
                break;
                
            case AcDb::kTextFit:
                hortzM = "Ajuster";
                break;
        }
        
        AcDb::TextVertMode vertMode = texte->verticalMode();
        
        switch( vertMode )
        {
            case AcDb::kTextBase:
                vertM = "Bas";
                break;
                
            case AcDb::kTextBottom:
                vertM = "Bas";
                break;
                
            case AcDb::kTextVertMid:
                vertM = "Centre";
                break;
                
            case AcDb::kTextTop:
                vertM = "Haut";
                break;
        }
        
        AcString s = vertM + " " + hortzM;
        
        if( DBFGetFieldIndex( dBFHandle, field ) )
            DBFWriteStringAttribute( dBFHandle, iShape, DBFGetFieldIndex( dBFHandle, field ), s );
            
    }
}

void writeRadius( AcDbEntity* ent, DBFHandle dBFHandle, int iShape, AcString fieldName )
{
    if( ent->isKindOf( AcDbCircle::desc() ) )
    {
        //Caster l'entité en cercle
        AcDbCircle* circle = AcDbCircle::cast( ent );
        
        //Recuperer le rayon du cercle
        auto radius = strToAcStr( to_string( circle->radius() ) );
        
        if( DBFGetFieldIndex( dBFHandle, fieldName ) )
            DBFWriteStringAttribute( dBFHandle, iShape, DBFGetFieldIndex( dBFHandle, fieldName ), radius );
    }
}

void writeTextStyle( AcDbEntity* ent, DBFHandle dBFHandle, int iShape, AcString field )
{
    if( ent->isKindOf( AcDbText::desc() ) )
    {
        AcDbText* txt = AcDbText::cast( ent );
        
        AcDbObjectId style = txt->textStyle();
        
        //Recuperer les txt style
        AcDbTextStyleTableRecord* txtStyle = NULL;
        Acad::ErrorStatus es = acdbOpenObject( txtStyle, style, AcDb::kForWrite );
        
        //Recuperer le nom de style
        AcString ac = txtStyle->getName();
        
        if( DBFGetFieldIndex( dBFHandle, field ) )
            DBFWriteStringAttribute( dBFHandle, iShape, DBFGetFieldIndex( dBFHandle, field ), ac );
            
        txtStyle->close();
    }
}

void writeText( AcDbEntity* ent, DBFHandle dBFHandle, int iShape, const vector<AcString>& fieldName )
{
    writeLine( ent, dBFHandle, iShape, fieldName );
    writeEntHandle( ent, dBFHandle, iShape, fieldName.at( 0 ) );
    writeLayerAttribute( ent, dBFHandle, iShape, fieldName.at( 1 ) );
    writeColorAttribute( ent, dBFHandle, iShape, fieldName.at( 2 ) );
    writeLynetypeAttribute( ent, dBFHandle, iShape, fieldName.at( 3 ) );
    writeLineTypeScaleAttribute( ent, dBFHandle, iShape, fieldName.at( 4 ) );
    writeLineWeightAttribute( ent, dBFHandle, iShape, fieldName.at( 5 ) );
    writeTransparency( ent, dBFHandle, iShape, fieldName.at( 6 ) );
    writeMaterialName( ent, dBFHandle, iShape, fieldName.at( 7 ) );
    writeXPos( ent, dBFHandle, iShape, fieldName.at( 8 ) );
    writeYPos( ent, dBFHandle, iShape, fieldName.at( 9 ) );
    writeZPos( ent, dBFHandle, iShape, fieldName.at( 10 ) );
    writeOrientation( ent, dBFHandle, iShape, fieldName.at( 11 ) );
    writeTextStr( ent, dBFHandle, iShape, fieldName.at( 12 ) );
    writeTexteHeight( ent, dBFHandle, iShape, fieldName.at( 13 ) );
    writeTextStyle( ent, dBFHandle, iShape, fieldName.at( 14 ) );
    writeTextAlign( ent, dBFHandle, iShape, fieldName.at( 15 ) );
    
}


void writeTextPoly( AcDbEntity* ent, AcDbEntity* entPoly, DBFHandle dBFHandle, int iShape, const vector<AcString>& fieldName )
{
    writeEntHandle( ent, dBFHandle, iShape, fieldName.at( 0 ) );
    writeLayerAttribute( ent, dBFHandle, iShape, fieldName.at( 1 ) );
    writeColorAttribute( ent, dBFHandle, iShape, fieldName.at( 2 ) );
    writeLynetypeAttribute( ent, dBFHandle, iShape, fieldName.at( 3 ) );
    writeLineTypeScaleAttribute( ent, dBFHandle, iShape, fieldName.at( 4 ) );
    writeLineWeightAttribute( ent, dBFHandle, iShape, fieldName.at( 5 ) );
    writeTransparency( ent, dBFHandle, iShape, fieldName.at( 6 ) );
    writeMaterialName( ent, dBFHandle, iShape, fieldName.at( 7 ) );
    writeXPos( ent, dBFHandle, iShape, fieldName.at( 8 ) );
    writeYPos( ent, dBFHandle, iShape, fieldName.at( 9 ) );
    writeZPos( ent, dBFHandle, iShape, fieldName.at( 10 ) );
    writeOrientation( ent, dBFHandle, iShape, fieldName.at( 11 ) );
    writeTextStr( ent, dBFHandle, iShape, strToAcStr( latin1_to_utf8( acStrToStr( fieldName.at( 12 ) ) ) ) );
    writeTexteHeight( ent, dBFHandle, iShape, fieldName.at( 13 ) );
    writeTextStyle( ent, dBFHandle, iShape, fieldName.at( 14 ) );
    writeTextAlign( ent, dBFHandle, iShape, fieldName.at( 15 ) );
    writeEntHandle( entPoly, dBFHandle, iShape, fieldName.at( 16 ) );
    writeLength2d( entPoly, dBFHandle, iShape, _T( "PLENGTH2D" ) );
    
    if( entPoly->isKindOf( AcDb3dPolyline::desc() ) )
        writeLength3d( entPoly, dBFHandle, iShape, _T( "PLENGTH3D" ) );
    else
        writeLength2d( entPoly, dBFHandle, iShape, _T( "PLENGTH3D" ) );
        
    writeArea( entPoly, dBFHandle, iShape, _T( "POLYAREA" ) );
    
    
}

void writePoint( AcDbEntity* ent, DBFHandle dBFHandle, const int& iShape, const vector<AcString>& fieldName )
{
    writeEntHandle( ent, dBFHandle, iShape, fieldName.at( 0 ) );
    writeLayerAttribute( ent, dBFHandle, iShape, fieldName.at( 1 ) );
    writeColorAttribute( ent, dBFHandle, iShape, fieldName.at( 2 ) );
    writeLynetypeAttribute( ent, dBFHandle, iShape, fieldName.at( 3 ) );
    writeLineTypeScaleAttribute( ent, dBFHandle, iShape, fieldName.at( 4 ) );
    writeLineWeightAttribute( ent, dBFHandle, iShape, fieldName.at( 5 ) );
    writeTransparency( ent, dBFHandle, iShape, fieldName.at( 6 ) );
    writeThicknessAttribute( ent, dBFHandle, iShape, fieldName.at( 7 ) );
    writeMaterialName( ent, dBFHandle, iShape, fieldName.at( 8 ) );
    writeXPos( ent, dBFHandle, iShape, fieldName.at( 9 ) );
    writeYPos( ent, dBFHandle, iShape, fieldName.at( 10 ) );
    writeZPos( ent, dBFHandle, iShape, fieldName.at( 11 ) );
    
}

void writeHatch( AcDbHatch* ent, DBFHandle dBFHandle, const int& iShape, const vector<AcString>& fieldName )
{
    writeEntHandle( ent, dBFHandle, iShape, fieldName.at( 0 ) );
    writeLayerAttribute( ent, dBFHandle, iShape, fieldName.at( 1 ) );
    writeColorAttribute( ent, dBFHandle, iShape, fieldName.at( 2 ) );
    writeLynetypeAttribute( ent, dBFHandle, iShape, fieldName.at( 3 ) );
    writeLineTypeScaleAttribute( ent, dBFHandle, iShape, fieldName.at( 4 ) );
    writeLineWeightAttribute( ent, dBFHandle, iShape, fieldName.at( 5 ) );
    writeTransparency( ent, dBFHandle, iShape, fieldName.at( 6 ) );
    writeMaterialName( ent, dBFHandle, iShape, fieldName.at( 7 ) );
    writeXPos( ent, dBFHandle, iShape, fieldName.at( 8 ) );
    writeYPos( ent, dBFHandle, iShape, fieldName.at( 9 ) );
    writeZPos( ent, dBFHandle, iShape, fieldName.at( 10 ) );
    writeArea( ent, dBFHandle, iShape, fieldName.at( 11 ) );
    writePattern( ent, dBFHandle, iShape, fieldName.at( 12 ) );
}

/* ==> End Utils function shape <== */

//Encode excel path into utf8 encodage
std::string latin1_to_utf8( const std::string& latin1 )
{
    std::string utf8;
    
    for( auto character : latin1 )
    {
        if( character >= 0 )
            utf8.push_back( character );
        else
        {
            utf8.push_back( 0xc0 | static_cast<std::uint8_t>( character ) >> 6 );
            utf8.push_back( 0x80 | ( static_cast<std::uint8_t>( character ) & 0x3f ) );
        }
    }
    
    return utf8;
}

void getEntityParams(
    const AcDbEntity* ent,
    AcString& entityLayertype,
    AcString& entityLineType,
    AcString& entityLayerName,
    AcString& entityColor,
    AcString& entityHandle,
    int& entityLineWeigth,
    AcString& entityLineWeigthStr,
    AcCmTransparency& entityTransp,
    AcString& entityTranspStr,
    double& entityLineTypeScale,
    AcString& entityPlotStyleName )
{
    /*
     *Les variables communs de l'entité
    */
    //Type de calque
    entityLayertype = ent->layer();
    
    //Type de ligne
    entityLineType = ent->linetype();
    
    //Nom du matériel
    entityLayerName = ent->material();
    
    //Couleur
    //Si la couleur est ByLayer
    if( ent->color().isByLayer() )
        entityColor = "ByLayer";
        
    //Si la couleur est ByBlock
    else if( ent->color().isByBlock() )
        entityColor = "ByBlock";
        
    //Si c'est pas l'un des deux
    else
        entityColor = numberToAcString( ent->color().red() ) + _T( "," ) + numberToAcString( ent->color().green() ) + _T( "," ) + numberToAcString( ent->color().blue() );
        
    //Taille de la ligne
    entityLineWeigth = ent->lineWeight();
    
    //Transparence
    entityTransp = ent->transparency();
    
    //LineType Scale
    entityLineTypeScale = ent->linetypeScale();
    
    //Condition d'affichage du LineWeight
    if( entityLineWeigth == -3 )
        entityLineWeigthStr = "Default";
    else if( entityLineWeigth == -2 )
        entityLineWeigthStr = "ByBlock";
    else if( entityLineWeigth == -1 )
        entityLineWeigthStr = "ByLayer";
    else
        entityLineWeigthStr = numberToAcString( entityLineWeigth );
        
    //Condition d'affichage du transparence
    if( entityTransp.isByBlock() )
        entityTranspStr = "ByBlock";
    else if( entityTransp.isByLayer() )
        entityTranspStr = "ByLayer";
    else if( entityTransp.isByAlpha() )
        entityTranspStr = numberToAcString( entityTransp.alpha() );
    else if( entityTransp.isSolid() )
        entityTranspStr = numberToAcString( entityTransp.alpha() );
        
    //Récupération du Handle de l'entité
    entityHandle = ent->handle().ascii();
    
    //Récupération du PlotStyleName
    entityPlotStyleName = ent->plotStyleName();
}



void exportPoly2d(
    AcDbPolyline*& pl2d,
    AcStringArray& pData )
{
    //--------------------------------------------<Variables>
    //Type de calque
    AcString poly2dLayerType,
             //Type de ligne
             poly2dLineType,
             //Matériel
             poly2dMaterialName,
             //Taille de ligne
             poly2dLineWeightStr,
             //Couleur
             poly2dColor,
             //Handle
             poly2dHandle,
             //Transparence
             poly2dTranspStr,
             //Status fermée/ouverte
             poly2dStatus,
             //Nom du style de tracé
             poly2dPlotStyleName;
             
    //Surface du polyline si elle est fermée
    double poly2dArea = 0;
    //Taille du ligne de la polyligne en int
    int poly2dLineWeigth;
    
    //Transparence de la polyliogne
    AcCmTransparency poly2dTransp;
    
    //Longueur, linetype scale
    double poly2dLength,
           poly2dLineTypeScale,
           poly2dThickness;
    //--------------------------------------------</Variables>
    
    
    //--------------------------------------------<Traitement>
    
    //Prendre les paramètres de l'entité
    getEntityParams(
        pl2d,
        poly2dLayerType,
        poly2dLineType,
        poly2dMaterialName,
        poly2dColor,
        poly2dHandle,
        poly2dLineWeigth,
        poly2dLineWeightStr,
        poly2dTransp,
        poly2dTranspStr,
        poly2dLineTypeScale,
        poly2dPlotStyleName );
        
    //Prendre la lonngueur de la polyligne
    poly2dLength = getLength( pl2d );
    
    //Vérification si la polyligne est fermée ou ouverte
    if( isClosed( pl2d ) )
    {
        poly2dStatus = "Oui";
        pl2d->getArea( poly2dArea );
    }
    
    else
        poly2dStatus = "Non";
        
    //Thickness de la polyligne 2d
    poly2dThickness = pl2d->thickness();
    string width_2d = to_string( pl2d->getConstantWidth() );
    
    pData.push_back( poly2dHandle );
    pData.push_back( poly2dLayerType );
    pData.push_back( poly2dColor );
    pData.push_back( poly2dLineType );
    pData.push_back( strToAcStr( to_string( poly2dLineTypeScale ) ) );
    pData.push_back( poly2dLineWeightStr );
    pData.push_back( poly2dTranspStr );
    pData.push_back( strToAcStr( to_string( poly2dThickness ) ) );
    pData.push_back( poly2dMaterialName );
    pData.push_back( strToAcStr( to_string( poly2dLength ) ) );
    pData.push_back( strToAcStr( width_2d ) );
    pData.push_back( poly2dStatus );
    pData.push_back( strToAcStr( to_string( poly2dArea ) ) );
    
    //--------------------------------------------</Traitement>
}


void exportPoly(
    AcDbEntity*& ent,
    AcStringArray& pData )
{
    //--------------------------------------------<Variables>
    //Type de calque
    AcString polyLayerType,
             //Type de ligne
             polyLineType,
             //Materiel
             polyMaterialName,
             //Taille de ligne
             polyLineWeightStr,
             //Couleur
             polyColor,
             //Handle
             polyHandle,
             //Transparence
             polyTranspStr,
             //Status fermée/ouverte
             polyStatus,
             //Style de tracé
             polyPlotStyleName;
             
    //Taille du ligne de la polyligne en int
    int polyLineWeigth;
    
    //Transparence de la polyliogne
    AcCmTransparency polyTransp;
    
    //Longueur, linetype scale
    double polyLength,
           polylength3d,
           polyLineTypeScale,
           polyThickness;
           
    bool has_thickness = true;
    
    //--------------------------------------------</Variables>
    
    
    //--------------------------------------------<Traitement>
    
    //Prendre les paramètres de l'entité
    getEntityParams(
        ent,
        polyLayerType,
        polyLineType,
        polyMaterialName,
        polyColor,
        polyHandle,
        polyLineWeigth,
        polyLineWeightStr,
        polyTransp,
        polyTranspStr,
        polyLineTypeScale,
        polyPlotStyleName );
        
    //verifier entity
    if( ent->isKindOf( AcDbPolyline::desc() ) )
    {
        AcDbPolyline* pl2d = AcDbPolyline::cast( ent );
        
        //Prendre la lonngueur de la polyligne
        polyLength = getLength( pl2d );
        polylength3d = polyLength;
        
        //Vérification si la polyligne est fermée ou ouverte
        if( isClosed( pl2d ) )
            polyStatus = "Oui";
        else
            polyStatus = "Non";
            
        //Thickness de la polyligne 2d
        polyThickness = pl2d->thickness();
        
        
    }
    
    else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
    {
        AcDb3dPolyline* pl3d = AcDb3dPolyline::cast( ent );
        
        //Prendre la lonngueur de la polyligne
        polyLength = get2DLengthPoly( pl3d );
        
        //Recuperer la longueur 3d de la polyline
        polylength3d = getLength( pl3d );
        
        //thickness
        has_thickness = false;
        
        //Vérification si la polyligne est fermée ou ouverte
        if( isClosed( pl3d ) )
            polyStatus = "Oui";
        else
            polyStatus = "Non";
            
    }
    
    pData.push_back( polyHandle );
    pData.push_back( polyLayerType );
    pData.push_back( polyColor );
    pData.push_back( polyLineType );
    pData.push_back( strToAcStr( to_string( polyLineTypeScale ) ) );
    pData.push_back( polyLineWeightStr );
    pData.push_back( polyTranspStr );
    
    if( has_thickness )
        pData.push_back( strToAcStr( to_string( polyThickness ) ) );
    else
        pData.push_back( _T( "N/A" ) );
        
    pData.push_back( polyMaterialName );
    
    pData.push_back( strToAcStr( to_string( polyLength ) ) );
    pData.push_back( strToAcStr( to_string( polylength3d ) ) );
    pData.push_back( polyStatus );
    
    
    //--------------------------------------------</Traitement>
}


void exportPoly3d(
    AcDb3dPolyline*& pl3d,
    AcStringArray& pData )
{

    //--------------------------------------------<Variables>
    
    //Type de calque
    AcString poly3dLayerType,
             //Type de ligne
             poly3dLineType,
             //Matériel
             poly3dMaterialName,
             //Couleur
             poly3dColor,
             //Handle
             poly3dHandle,
             //Taille de ligne en string
             poly3dLineWeightStr,
             //Transparence
             poly3dTranspStr,
             //Status fermée/ouverte
             polyStatus,
             //Nom  du style de tracé
             poly3dPlotStyleName;
             
    //Longueur
    double poly3dLength,
           //Linetype scale
           poly3dLineTypeScale,
           //Longueur 2D
           poly3dLength2d,
           //Surface de la polyline 3d
           poly3dArea
           ;
           
    poly3dArea = 0;
    //Taille du ligne de la polyligne en int
    int poly3dLineWeigth;
    
    //Transparence de la polyligne
    AcCmTransparency poly3dTransp;
    
    //--------------------------------------------</Variable>
    
    //--------------------------------------------<Traitement>
    
    //Prendre les paramètres de l'entité
    getEntityParams(
        pl3d,
        poly3dLayerType,
        poly3dLineType,
        poly3dMaterialName,
        poly3dColor,
        poly3dHandle,
        poly3dLineWeigth,
        poly3dLineWeightStr,
        poly3dTransp,
        poly3dTranspStr,
        poly3dLineTypeScale,
        poly3dPlotStyleName );
        
    //Prendre la longueur de la polyligne
    poly3dLength = getLengthPoly( pl3d );
    
    //Prendre la longueur 2d de la polyligne
    poly3dLength2d = get2DLengthPoly( pl3d );
    
    //Verification si fermée ou ouverte
    if( isClosed( pl3d ) )
    {
        polyStatus = "Oui";
        pl3d->getArea( poly3dArea );
    }
    
    else
        polyStatus = "Non";
        
    // Mettre tous les données dans le tableau
    pData.push_back( poly3dHandle );
    pData.push_back( poly3dLayerType );
    pData.push_back( poly3dColor );
    pData.push_back( poly3dLineType );
    pData.push_back( strToAcStr( to_string( poly3dLineTypeScale ) ) );
    pData.push_back( poly3dLineWeightStr );
    pData.push_back( poly3dTranspStr );
    pData.push_back( poly3dMaterialName );
    pData.push_back( strToAcStr( to_string( poly3dLength2d ) ) );
    pData.push_back( strToAcStr( to_string( poly3dLength ) ) );
    pData.push_back( polyStatus );
    pData.push_back( strToAcStr( to_string( poly3dArea ) ) );
    
    //--------------------------------------------</Traitement>
}


void exportClosedPoly(
    AcDbEntity*& ent,
    AcStringArray& pData )
{

    //--------------------------------------------<Variables>
    
    //Type de calque
    AcString poly3dLayerType,
             //Type de ligne
             poly3dLineType,
             //Matériel
             poly3dMaterialName,
             //Couleur
             poly3dColor,
             //Handle
             poly3dHandle,
             //Taille de ligne en string
             poly3dLineWeightStr,
             //Transparence
             poly3dTranspStr,
             //Status fermée/ouverte
             polyStatus,
             //Nom  du style de tracé
             poly3dPlotStyleName;
             
    //Longueur
    double poly3dLength,
           //Linetype scale
           poly3dLineTypeScale,
           //Longueur 2D
           poly3dLength2d,
           //Surface de la polyline 3d
           poly3dArea
           ;
           
    //Declarer l'épaisseur
    double entThickness;
    
    poly3dArea = 0;
    //Taille du ligne de la polyligne en int
    int poly3dLineWeigth;
    
    //Transparence de la polyligne
    AcCmTransparency poly3dTransp;
    
    //--------------------------------------------</Variable>
    
    //--------------------------------------------<Traitement>
    
    //Prendre les paramètres de l'entité
    getEntityParams(
        ent,
        poly3dLayerType,
        poly3dLineType,
        poly3dMaterialName,
        poly3dColor,
        poly3dHandle,
        poly3dLineWeigth,
        poly3dLineWeightStr,
        poly3dTransp,
        poly3dTranspStr,
        poly3dLineTypeScale,
        poly3dPlotStyleName );
        
    if( ent->isKindOf( AcDbPolyline::desc() ) )
    {
        //Caster l'entité
        AcDbPolyline* p2d = AcDbPolyline::cast( ent );
        
        //Prendre la longueur de la polyligne
        poly3dLength2d = getLength( p2d );
        
        //Recuperer la surface
        p2d->getArea( poly3dArea );
        
        entThickness = p2d->thickness();
        
    }
    
    else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
    {
        //Caster l'entité
        AcDb3dPolyline* p3d = AcDb3dPolyline::cast( ent );
        
        //Prendre la longueur de la polyligne
        poly3dLength = getLengthPoly( p3d );
        
        //Prendre la longueur 2d de la polyligne
        poly3dLength2d = get2DLengthPoly( p3d );
        
        //Recuperer la surface
        p3d->getArea( poly3dArea );
        
        entThickness = 0;
    }
    
    
    // Mettre tous les données dans le tableau
    pData.push_back( poly3dHandle );
    pData.push_back( poly3dLayerType );
    pData.push_back( poly3dColor );
    pData.push_back( poly3dLineType );
    pData.push_back( strToAcStr( to_string( poly3dLineTypeScale ) ) );
    pData.push_back( poly3dLineWeightStr );
    pData.push_back( poly3dTranspStr );
    
    if( ent->isKindOf( AcDbPolyline::desc() ) )
        pData.push_back( strToAcStr( to_string( entThickness ) ) );
    else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
        pData.push_back( strToAcStr( "N/A" ) );
        
    pData.push_back( poly3dMaterialName );
    pData.push_back( strToAcStr( to_string( poly3dLength2d ) ) );
    
    if( ent->isKindOf( AcDbPolyline::desc() ) )
        pData.push_back( strToAcStr( to_string( poly3dLength2d ) ) );
    else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
        pData.push_back( strToAcStr( to_string( poly3dLength ) ) );
        
    pData.push_back( strToAcStr( to_string( poly3dArea ) ) );
    
    //--------------------------------------------</Traitement>
}


void exportPoint(
    AcDbPoint*& pt,
    AcStringArray& pData )
{
    //Type de calque
    AcString pointLayertype,
             //Type de ligne
             pointLineType,
             //Nom du matériel
             pointMaterialName,
             //Couleur du point
             pointColor,
             //Handle du point
             pointHandle,
             //Transparence du point en string
             pointTranspStr,
             //Taille de ligne en String
             pointLineWeigthStr,
             //Nom du style de tracé
             pointPlotStyleName;
             
    //Taille de ligne en int
    int pointLineWeigth;
    
    //Echelle de type de ligne
    double pointLineTypeScale,
           //Thickness
           pointThickness;
           
    //La position en X, Y, Z du point
    float pointPositionX,
          pointPositionY,
          pointPositionZ;
          
    //Transparence du point
    AcCmTransparency pointTransp;
    
    //Prendre les paramètres de l'entité
    getEntityParams(
        pt,
        pointLayertype,
        pointLineType,
        pointMaterialName,
        pointColor,
        pointHandle,
        pointLineWeigth,
        pointLineWeigthStr,
        pointTransp,
        pointTranspStr,
        pointLineTypeScale,
        pointPlotStyleName );
        
    //Thickness de/des point(s)
    pointThickness = pt->thickness();
    
    //Position sur les axes
    pointPositionX = pt->position().x;
    pointPositionY = pt->position().y;
    pointPositionZ = pt->position().z;
    
    // Mettre les données dans le tableau
    pData.push_back( pointHandle );
    pData.push_back( pointLayertype );
    pData.push_back( pointColor );
    pData.push_back( pointLineType );
    pData.push_back( strToAcStr( to_string( pointLineTypeScale ) ) );
    pData.push_back( pointLineWeigthStr );
    pData.push_back( pointTranspStr );
    pData.push_back( strToAcStr( to_string( pointThickness ) ) );
    pData.push_back( pointMaterialName );
    pData.push_back( strToAcStr( to_string( pointPositionX ) ) );
    pData.push_back( strToAcStr( to_string( pointPositionY ) ) );
    pData.push_back( strToAcStr( to_string( pointPositionZ ) ) );
}



void exportEntity(
    AcDbEntity*& ent,
    AcStringArray& pData )
{
    //Type de calque
    AcString entLayertype,
             //Type de ligne
             entLineType,
             //Nom du matériel
             entMaterialName,
             //Couleur du point
             entColor,
             //Handle du point
             entHandle,
             //Transparence du point en string
             entTranspStr,
             //Taille de ligne en String
             entLineWeigthStr,
             //Nom du style de tracé
             entPlotStyleName;
             
    //Taille de ligne en int
    int entLineWeigth;
    
    //Echelle de type de ligne
    double entLineTypeScale,
           //Thickness
           entThickness;
           
    //Transparence du point
    AcCmTransparency entTransp;
    
    //bool thickness
    bool has_thickness = true;
    bool has_position = false;
    
    //Declarer la position de l'entité
    AcGePoint3d pos_entity = AcGePoint3d::kOrigin;
    
    //Prendre les paramètres de l'entité
    getEntityParams(
        ent,
        entLayertype,
        entLineType,
        entMaterialName,
        entColor,
        entHandle,
        entLineWeigth,
        entLineWeigthStr,
        entTransp,
        entTranspStr,
        entLineTypeScale,
        entPlotStyleName );
        
    ////Thickness de/des point(s)
    if( ent->isKindOf( AcDbLine::desc() ) )
    {
        //Caster l'entité
        AcDbLine* ln = AcDbLine::cast( ent );
        
        //Recuperer le thickness
        entThickness = ln->thickness();
    }
    
    else if( ent->isKindOf( AcDbPolyline::desc() ) )
    {
        //Caster l'entité
        AcDbPolyline* ln = AcDbPolyline::cast( ent );
        
        //Recuperer le thickness
        entThickness = ln->thickness();
    }
    
    else if( ent->isKindOf( AcDbPoint::desc() ) )
    {
        //Caster l'entité
        AcDbPoint* ln = AcDbPoint::cast( ent );
        
        //Recuperer le thickness
        entThickness = ln->thickness();
        
        //Mettre has position à true
        has_position = true;
        
        //Recuperer la position de l'entité
        pos_entity = ln->position();
    }
    
    else if( ent->isKindOf( AcDbText::desc() ) )
    {
        //Caster l'entité
        AcDbText* ln = AcDbText::cast( ent );
        
        //Recuperer le thickness
        entThickness = ln->thickness();
        
        //Mettre has position à true
        has_position = true;
        
        //Recuperer la position de l'entité
        pos_entity = ln->position();
    }
    
    else if( ent->isKindOf( AcDbCircle::desc() ) )
    {
        //Caster l'entité
        AcDbCircle* ln = AcDbCircle::cast( ent );
        
        //Recuperer le thickness
        entThickness = ln->thickness();
        
        //Mettre has position à true
        has_position = true;
        
        //Recuperer la position de l'entité
        pos_entity = ln->center();
    }
    
    else if( ent->isKindOf( AcDbBlockReference::desc() ) )
    {
        //Caster l'entité
        AcDbBlockReference* ln = AcDbBlockReference::cast( ent );
        
        //Mettre has_thickness à false
        has_thickness = false;
        
        //Mettre has position à true
        has_position = true;
        
        //Recuperer la position de l'entité
        pos_entity = ln->position();
    }
    
    else if( ent->isKindOf( AcDbMText::desc() ) )
    {
        //Caster l'entité
        AcDbMText* ln = AcDbMText::cast( ent );
        
        //Mettre has_thickness à false
        has_thickness = false;
        
        //Mettre has position à true
        has_position = true;
        
        //Recuperer la position de l'entité
        pos_entity = ln->location();
    }
    
    else if( ent->isKindOf( AcDbEllipse::desc() ) )
    {
        //Caster l'entité
        AcDbEllipse* ln = AcDbEllipse::cast( ent );
        
        //Mettre has_thickness à false
        has_thickness = false;
        
        //Mettre has position à true
        has_position = true;
        
        //Recuperer la position de l'entité
        pos_entity = ln->center();
    }
    
    // Mettre les données dans le tableau
    pData.push_back( entHandle );
    pData.push_back( entLayertype );
    pData.push_back( entColor );
    pData.push_back( entLineType );
    pData.push_back( strToAcStr( to_string( entLineTypeScale ) ) );
    pData.push_back( entLineWeigthStr );
    pData.push_back( entTranspStr );
    
    if( has_thickness )
        pData.push_back( strToAcStr( to_string( entThickness ) ) );
    else
        pData.push_back( strToAcStr( "N/A" ) );
        
    pData.push_back( entMaterialName );
    
    if( has_position )
    {
        pData.push_back( strToAcStr( to_string( pos_entity.x ) ) );
        pData.push_back( strToAcStr( to_string( pos_entity.y ) ) );
        pData.push_back( strToAcStr( to_string( pos_entity.z ) ) );
    }
    
    else
    {
        pData.push_back( strToAcStr( "N/A" ) );
        pData.push_back( strToAcStr( "N/A" ) );
        pData.push_back( strToAcStr( "N/A" ) );
    }
}

void exportCurve(
    AcDbEntity*& ent,
    AcStringArray& pData )
{
    //Type de calque
    AcString entLayertype,
             //Type de ligne
             entLineType,
             //Nom du matériel
             entMaterialName,
             //Couleur du point
             entColor,
             //Handle du point
             entHandle,
             //Transparence du point en string
             entTranspStr,
             //Taille de ligne en String
             entLineWeigthStr,
             //Nom du style de tracé
             entPlotStyleName;
             
    //Taille de ligne en int
    int entLineWeigth;
    
    //Echelle de type de ligne
    double entLineTypeScale,
           //Thickness
           entThickness;
           
    //Transparence du point
    AcCmTransparency entTransp;
    
    //bool thickness
    bool has_thickness = true;
    bool has_length = true;
    
    //Les longueurs
    double length_2d;
    double length_3d;
    AcString status;
    
    //Prendre les paramètres de l'entité
    getEntityParams(
        ent,
        entLayertype,
        entLineType,
        entMaterialName,
        entColor,
        entHandle,
        entLineWeigth,
        entLineWeigthStr,
        entTransp,
        entTranspStr,
        entLineTypeScale,
        entPlotStyleName );
        
    ////Thickness de/des point(s)
    if( ent->isKindOf( AcDbLine::desc() ) )
    {
        //Caster l'entité
        AcDbLine* ln = AcDbLine::cast( ent );
        
        //Recuperer le thickness
        entThickness = ln->thickness();
        
        //Longueur
        length_2d = getDistance2d( ln->startPoint(), ln->endPoint() );
        
        length_3d = length_2d;
        
        //Statut
        status = _T( "N/A" );
    }
    
    else if( ent->isKindOf( AcDbPolyline::desc() ) )
    {
        //Caster l'entité
        AcDbPolyline* ln = AcDbPolyline::cast( ent );
        
        length_2d = getLength( ln );
        
        length_3d = length_2d;
        
        //Recuperer le thickness
        entThickness = ln->thickness();
        
        if( isClosed( ln ) )
            status = _T( "Oui" );
        else
            status = _T( "Non" );
    }
    
    else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
    {
        //Caster l'entité
        AcDb3dPolyline* ln = AcDb3dPolyline::cast( ent );
        
        has_thickness = false;
        
        length_2d = get2DLengthPoly( ln );
        
        length_3d = getLength( ln );
        
        if( isClosed( ln ) )
            status = _T( "Oui" );
        else
            status = _T( "Non" );
            
    }
    
    else if( ent->isKindOf( AcDbCircle::desc() ) )
    {
        //Caster l'entité
        AcDbCircle* ln = AcDbCircle::cast( ent );
        
        //Recuperer le thickness
        entThickness = ln->thickness();
        
        has_length = false;
        
        status = _T( "N/A" );
    }
    
    else if( ent->isKindOf( AcDbEllipse::desc() ) )
    {
        //Caster l'entité
        AcDbEllipse* ln = AcDbEllipse::cast( ent );
        
        //Mettre has_thickness à false
        has_thickness = false;
        
        has_length = false;
        
        status = _T( "N/A" );
        
    }
    
    // Mettre les données dans le tableau
    pData.push_back( entHandle );
    pData.push_back( entLayertype );
    pData.push_back( entColor );
    pData.push_back( entLineType );
    pData.push_back( strToAcStr( to_string( entLineTypeScale ) ) );
    pData.push_back( entLineWeigthStr );
    pData.push_back( entTranspStr );
    
    if( has_thickness )
        pData.push_back( strToAcStr( to_string( entThickness ) ) );
    else
        pData.push_back( strToAcStr( "N/A" ) );
        
    pData.push_back( entMaterialName );
    
    if( has_length )
    {
        pData.push_back( strToAcStr( to_string( length_2d ) ) );
        pData.push_back( strToAcStr( to_string( length_3d ) ) );
    }
    
    else
    {
        pData.push_back( strToAcStr( "N/A" ) );
        pData.push_back( strToAcStr( "N/A" ) );
    }
    
    pData.push_back( status );
    
}


void exportBlock(
    AcDbBlockReference*& blk,
    AcStringArray& pData )
{

    //------------------------------------------------------<Variables>
    
    //Nom du bloc
    AcString blockName,
             //Type de calque
             blockLayerType,
             //Type de ligne
             blockLineType,
             //Matériel
             blockMaterialName,
             //Couleur
             blockColor,
             //Handle
             blockHandle,
             //Taille de ligne en string
             blockLineWeigthStr,
             //Transparence
             blockTranspStr,
             //Nom du style de tracé
             blockPlotStyleName,
             //Liste des attribut
             Attr;
             
    //Echelle de type de ligne
    double blockLineTypeScale;
    
    //Taille du block en int
    int blockLineWeigth;
    
    //Transparence du block
    AcCmTransparency blockTransp;
    
    //------------------------------------------------------</Variables>
    
    //------------------------------------------------------<Traitements>
    
    //Les paramètres de l'entité block
    getEntityParams(
        blk,
        blockLayerType,
        blockLineType,
        blockMaterialName,
        blockColor,
        blockHandle,
        blockLineWeigth,
        blockLineWeigthStr,
        blockTransp,
        blockTranspStr,
        blockLineTypeScale,
        blockPlotStyleName
    );
    
    //Le nom du block
    getBlockName( blockName, blk );
    
    //Position du block sur les axes
    float blockPositionX = blk->position().x;
    float blockPositionY = blk->position().y;
    float blockPositionZ = blk->position().z;
    
    //Rotation du block
    string rotation = to_string( blk->rotation() );
    auto scale = blk->scaleFactors();
    
    
    // Mettre les données dans le tableaux
    pData.push_back( blockHandle );
    pData.push_back( blockLayerType );
    pData.push_back( blockName );
    pData.push_back( blockColor );
    pData.push_back( blockLineType );
    pData.push_back( strToAcStr( to_string( blockLineTypeScale ) ) );
    pData.push_back( blockLineWeigthStr );
    pData.push_back( blockTranspStr );
    pData.push_back( blockMaterialName );
    pData.push_back( strToAcStr( to_string( blockPositionX ) ) );
    pData.push_back( strToAcStr( to_string( blockPositionY ) ) );
    pData.push_back( strToAcStr( to_string( blockPositionZ ) ) );
    pData.push_back( strToAcStr( rotation ) );
    
    
    //------------------------------------------------------</Traitements>
}



void exportText(
    AcDbText*& txt,
    AcStringArray& pData )
{
    //------------------------------------------------------<Variables>
    
    //Type de calque
    AcString txtLayerType,
             //Type de ligne
             txtLineType,
             //Matériel
             txtMaterialName,
             //Couleur
             txtColor,
             //Handle
             txtHandle,
             //Taille de ligne en string
             txtLineWeigthStr,
             //Transparence
             txtTranspStr,
             //Nom du style de tracé
             txtPlotStyleName,
             //Contenu du texte
             txtContent,
             // Style du texte
             txtStyle;
             
    //Echelle de type de ligne
    double txtLineTypeScale;
    
    //Taille du block en int
    int txtLineWeigth;
    
    //Transparence du block
    AcCmTransparency txtTransp;
    
    //------------------------------------------------------</Variables>
    
    //------------------------------------------------------<Traitements>
    
    // Prendre les valeurs des propriétes communs
    getEntityParams(
        txt,
        txtLayerType,
        txtLineType,
        txtMaterialName,
        txtColor,
        txtHandle,
        txtLineWeigth,
        txtLineWeigthStr,
        txtTransp,
        txtTranspStr,
        txtLineTypeScale,
        txtPlotStyleName
    );
    
    // Position du texte sur le axes
    float txtPositionX = txt->position().x,
          txtPositionY = txt->position().y,
          txtPositionZ = txt->position().z,
          
          // Alignement du texte sur l'axe XYZ
          txtAlignmentX = txt->alignmentPoint().x,
          txtAlignmentY = txt->alignmentPoint().y,
          txtAlignmentZ = txt->alignmentPoint().z,
          
          // Prendre la taille du texte
          txtHeight = txt->height(),
          
          // Prendre le width factor du texte
          txtWidthFactor = txt->widthFactor(),
          
          // La rotation du texte
          txtRotation = txt->rotation(),
          
          // L'angle oblique du texte
          txtOblique = txt->oblique();
          
    // Prendre le style du texte en AcString
    getStyleStringName( txtStyle, txt->textStyle() );
    
    // Contenu du texte
    txtContent = txt->textString();
    
    // Thickness du texte
    double txtThickness = txt->thickness();
    
    
    // Alignement à l'horizontale
    AcDb::TextHorzMode hortzMode = txt->horizontalMode();
    
    // Le variables pour les résultats de l'alignement
    AcString hortzM, vertM;
    
    switch( hortzMode )
    {
        case AcDb::kTextLeft:
            hortzM = "Gauche";
            break;
            
        case AcDb::kTextCenter:
            hortzM = "Centre";
            break;
            
        case AcDb::kTextRight:
            hortzM = "Droite";
            break;
            
        case AcDb::kTextAlign:
            hortzM = "Aligne";
            break;
            
        case AcDb::kTextMid:
            hortzM = "Milieu";
            break;
            
        case AcDb::kTextFit:
            hortzM = "Ajuster";
            break;
    }
    
    // Alignement à la verticale
    AcDb::TextVertMode vertMode = txt->verticalMode();
    
    switch( vertMode )
    {
        case AcDb::kTextBase:
            vertM = "Bas";
            break;
            
        case AcDb::kTextBottom:
            vertM = "Bas";
            break;
            
        case AcDb::kTextVertMid:
            vertM = "Centre";
            break;
            
        case AcDb::kTextTop:
            vertM = "Haut";
            break;
    }
    
    //------------------------------------------------------</Traitements>
    
    // Insertion des données dans le tableau
    pData.push_back( txtHandle );
    pData.push_back( txtLayerType );
    pData.push_back( txtColor );
    pData.push_back( txtLineType );
    pData.push_back( strToAcStr( to_string( txtLineTypeScale ) ) );
    pData.push_back( txtLineWeigthStr );
    pData.push_back( txtTranspStr );
    pData.push_back( strToAcStr( to_string( txtThickness ) ) );
    pData.push_back( txtMaterialName );
    pData.push_back( strToAcStr( to_string( txtPositionX ) ) );
    pData.push_back( strToAcStr( to_string( txtPositionY ) ) );
    pData.push_back( strToAcStr( to_string( txtPositionZ ) ) );
    pData.push_back( strToAcStr( to_string( txtRotation ) ) );
    pData.push_back( txtContent );
    pData.push_back( strToAcStr( to_string( txtHeight ) ) );
    /*  pData.push_back( strToAcStr( to_string( txtAlignmentX ) ) );
      pData.push_back( strToAcStr( to_string( txtAlignmentY ) ) );
      pData.push_back( strToAcStr( to_string( txtAlignmentZ ) ) );*/
    pData.push_back( txtStyle );
    pData.push_back( vertM );
    pData.push_back( hortzM );
    /*pData.push_back( strToAcStr( to_string( txtWidthFactor ) ) );
    pData.push_back( strToAcStr( to_string( txtOblique ) ) );*/
}



void exportMText(
    AcDbMText*& mTxt,
    AcStringArray& pData )
{
    //------------------------------------------------------<Variables>
    
    //Type de calque
    AcString mTxtLayerType,
             //Type de ligne
             mTxtLineType,
             //Matériel
             mTxtMaterialName,
             //Couleur
             mTxtColor,
             //Handle
             mTxtHandle,
             //Taille de ligne en string
             mTxtLineWeigthStr,
             //Transparence
             mTxtTranspStr,
             //Nom du style de tracé
             mTxtPlotStyleName,
             //Contenu du texte multiple
             mTxtContent;
             
    //Echelle de type de ligne
    double mTxtLineTypeScale;
    
    //Taille du block en int
    int mTxtLineWeigth;
    
    //Transparence du block
    AcCmTransparency mTxtTransp;
    
    //------------------------------------------------------</Variables>
    
    //------------------------------------------------------<Traitements>
    
    // Prendre les paramètres par défaut du texte multiple
    getEntityParams(
        mTxt,
        mTxtLayerType,
        mTxtLineType,
        mTxtMaterialName,
        mTxtColor,
        mTxtHandle,
        mTxtLineWeigth,
        mTxtLineWeigthStr,
        mTxtTransp,
        mTxtTranspStr,
        mTxtLineTypeScale,
        mTxtPlotStyleName
    );
    
    // Style du texte
    AcString mTxtStyle;
    
    // Prendre le style du texte en AcString
    getStyleStringName( mTxtStyle, mTxt->textStyle() );
    
    // Prendre la taille du mTexte
    float mTxtHeight = mTxt->height();
    
    // Prendre la largeur du mText
    float mTxtWidth = mTxt->width();
    
    // Taille du texte du mText
    float mTxtTextHeight = mTxt->textHeight();
    
    // Position sur mText sur les axes
    float mTextPosX = mTxt->location().x,   // Axe X
          mTextPosY = mTxt->location().y,   // Axe Y
          mTextPosZ = mTxt->location().z;   // Axe Z
          
          
    // ----------------------------------Flow direction
    AcString flDir = "";
    
    // -> Vérification du flow direction
    if( mTxt->flowDirection() == 1 )
        flDir = "Gauche à droite";
    else if( mTxt->flowDirection() == 2 )
        flDir = "Droite à gauche";
    else if( mTxt->flowDirection() == 3 )
        flDir = "Haut en bat";
    else if( mTxt->flowDirection() == 4 )
        flDir = "Bas en  haut";
    else if( mTxt->flowDirection() == 5 )
        flDir = "ByStyle";
        
    // La rotation du texte
    double txtRotation = mTxt->rotation();
    
    // LineSpaceFactor du mText
    double mTxtLineSpaceFactor = mTxt->lineSpacingFactor();
    
    // LineSpaceStyle du mText en AcString
    AcString mTxtLineSpaceStyle = "";
    
    // Verfication de la valeur
    if( mTxt->lineSpacingStyle() == 1 )
    {
        // La valeur minimum
        mTxtLineSpaceStyle = "AtLeast";
    }
    
    else if( mTxt->lineSpacingStyle() == 2 )
    {
        // La valeur absolue
        mTxtLineSpaceStyle = "Exactly";
    }
    
    // Contenu du texte multiple
    mTxtContent = mTxt->contents();
    
    // Point d'insertion du mText suivant les axes
    float insertPtX = mTxt->location().x;
    float insertPtY = mTxt->location().y;
    float insertPtZ = mTxt->location().z;
    
    // Affichage du point d'insertion
    AcString insertPt = _T( "Point : ( " ) + numberToAcString( insertPtX ) + _T( ", " ) + numberToAcString( insertPtY ) + _T( ", " ) + numberToAcString( insertPtZ ) + _T( " )" );
    
    //------------------------------------------------------</Traitements>
    
    // Mettre les données dans le tableau
    pData.push_back( mTxtHandle );
    pData.push_back( mTxtLayerType );
    pData.push_back( mTxtColor );
    pData.push_back( mTxtLineType );
    pData.push_back( strToAcStr( to_string( mTxtLineTypeScale ) ) );
    pData.push_back( mTxtLineWeigthStr );
    pData.push_back( mTxtTranspStr );
    pData.push_back( strToAcStr( to_string( mTxtTextHeight ) ) );
    pData.push_back( mTxtMaterialName );
    pData.push_back( insertPt );
    pData.push_back( strToAcStr( to_string( txtRotation ) ) );
    pData.push_back( mTxtContent );
    pData.push_back( strToAcStr( to_string( mTxtHeight ) ) );
    pData.push_back( mTxtStyle );
    //pData.push_back( strToAcStr( to_string( mTxtWidth ) ) );
    pData.push_back( strToAcStr( to_string( mTxtLineSpaceFactor ) ) );
    pData.push_back( mTxtLineSpaceStyle );
    pData.push_back( flDir );
}




bool writeXmlHeader( ofstream& xmlPath )
{
    xmlPath << "<?xml version=\"1.0\"?>\n<LandXML\n	xmlns=\"http://www.landxml.org/schema/LandXML-1.2\"\n	xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"\n	xsi:schemaLocation=\"http://www.landxml.org/schema/LandXML-1.2 http://www.landxml.org/schema/LandXML-1.2/LandXML-1.2.xsd\"\n	date=\"2021-05-28\"\n	time=\"09:48:21\"\n	version=\"1.2\"\n	language=\"English\"\n	readOnly=\"false\">\n	<Units>\n		<Metric\n			areaUnit=\"squareMeter\"\n			linearUnit=\"meter\"\n			volumeUnit=\"cubicMeter\"\n			temperatureUnit=\"celsius\"\n			pressureUnit=\"milliBars\"\n			diameterUnit=\"millimeter\"\n			angularUnit=\"decimal degrees\"\n			directionUnit=\"decimal degrees\">\n		</Metric>\n	</Units>\n	<Alignments name=\"\">\n\n";
    return true;
}

bool writeXmlFooter( ofstream& xmlPath )
{
    xmlPath << "	</Alignments>\n</LandXML>\n";
    return true;
}

vector<LinePoly> poly3dToSeg( AcDb3dPolyline* poly3d )
{
    //Resultat
    vector<LinePoly> vecRes;
    
    //Recuperer les sommets de la polyligne
    AcGePoint3dArray vertexes;
    vector<double> xPos;
    getVertexesPoly( poly3d, vertexes, xPos );
    
    //Recuperer la taille du vertexes
    long sizeVertexes = vertexes.size();
    
    AcGePoint3d ptStart = vertexes[0];
    
    //Boucle pour creer les segments
    for( long i = 0; i < sizeVertexes - 1; i++ )
    {
        //Creer un line poly
        LinePoly seg;
        
        //Ajouter le premier points
        seg.ptStart = ptStart;
        
        //Recuperer le second points
        AcGePoint3d ptEnd = vertexes[i + 1];
        
        //Ajouter le second points
        seg.ptEnd = ptEnd;
        
        //Recuperer la distance 2d entre les deux points
        double distance = getDistance2d( ptStart, ptEnd );
        
        //Setter la distance
        seg.curveDistance = distance;
        
        //Changer le point
        ptStart = ptEnd;
        
        //Ajouter le segment dans le vecteur
        vecRes.push_back( seg );
    }
    
    //Retourner le resultat
    return vecRes;
}


bool exportPoly3dToXml( ofstream& xmlPath,
    const vector<LinePoly>& segPoly,
    const AcString& layerPoly,
    const int& laySize,
    const int& i )
{
    //Recuperer la taille du vecteur de segment
    int sizeVec = segPoly.size();
    
    //Recuperer le name
    AcString name = layerPoly;
    
    if( laySize > 1 )
        name = layerPoly + numberToAcString( i + 1 );
        
    //Ajouter l'entete dans le xml
    xmlPath << "		<Alignment name=\"" << name << "\" staStart=\"0.00\" desc=\"Polyligne 3D\">\n			<CoordGeom>\n";
    
    //Boucle sur les segments
    for( int i = 0; i < sizeVec; i++ )
    {
        //Recuperer le ieme segment
        LinePoly seg = segPoly[i];
        
        //Recuperer le premier point
        double ptStartX = seg.ptStart.x;
        double ptStartY = seg.ptStart.y;
        
        //Recuperer le dernier point
        double ptEndX = seg.ptEnd.x;
        double ptEndY = seg.ptEnd.y;
        
        //Ecrire les informations dans le xml
        xmlPath << "				<Line>\n					<Start>" << std::fixed << std::setprecision( 12 ) << ptStartY << " " << ptStartX << "</Start>\n";
        xmlPath << "					<End>" << std::fixed << std::setprecision( 12 ) << ptEndY << " " << ptEndX << "</End>\n				</Line>\n";
    }
    
    xmlPath << "			</CoordGeom>\n			<Profile name=\"" << name << "\">\n				<ProfAlign name=\"" << name << "\">\n";
    
    double distCurv = 0;
    
    //Reboucle sur les segments
    for( int i = 0; i < sizeVec; i++ )
    {
        //Recuperer le ieme segment
        LinePoly seg = segPoly[i];
        
        //Recuperer la distance cuviligne
        if( i == 0 )
            distCurv = 0;
        else
        {
            LinePoly seg0 = segPoly[i - 1];
            distCurv += seg0.curveDistance;
        }
        
        //Recuperer le z du premier sommet
        double z = seg.ptStart.z;
        
        //Ecrire les informations dans le cml
        xmlPath << "					<PVI>" << std::fixed << std::setprecision( 12 ) << distCurv << " " << z << "</PVI>\n";
    }
    
    //Ecrire le dernier point
    xmlPath << "					<PVI>" << std::fixed << std::setprecision( 12 ) << distCurv + segPoly[sizeVec - 1].curveDistance << " " << segPoly[sizeVec - 1].ptEnd.z << "</PVI>\n";
    
    //Ecrire le reste du xml
    xmlPath << "				</ProfAlign>\n			</Profile>\n		</Alignment>\n";
    
    return true;
}


int exportFaceToXml(
    ofstream* file,
    string& sLayer )
{
    // on demande la sélection
    ads_name faceSelection;
    //acutPrintf( _T( "\n Sélectionner les faces \n" ) );
    long length = getSelectionSet( faceSelection, "X", "3DFACE", sLayer.c_str() );
    
    // récupère le nombre d'objets séletionnés
    if( ( acedSSLength( faceSelection, &length ) != RTNORM ) || ( length == 0 ) )
    {
        //acutPrintf( _T( " Selection failed \n" ) );
        acedSSFree( faceSelection );
        return 0;
    }
    
    //Completer le nom du calque
    *file << "\n\t<Surface name=\"" + sLayer + "\" desc=\"Description\">";
    *file << "\n\t\t<Definition surfType=\"TIN\">";
    *file << "\n\t\t\t<Pnts>";
    
    // vector des positions de la sélection
    vector< AcGePoint3d > positions;
    positions.resize( 0 );
    
    // vecteur avec les index correspondants aux positions des sommets
    vector< vector< int >> index;
    
    //vecteur avec les index de la face i
    vector< int > indexFaceI;
    index.resize( 0 );
    
    ads_name ent;
    AcDbObjectId id = AcDbObjectId::kNull;
    
    /// 1. On parcourt toutes les faces 3D
    
    vector< AcGePoint3d >::iterator it;
    
    // on boucle sur les faces
    for( int i = 0; i < length; i++ )
    {
    
        if( acedSSName( faceSelection, i, ent ) != RTNORM )
            continue;
            
        if( acdbGetObjectId( id, ent ) != Acad::eOk )
            continue;
            
        AcDbEntity* pEnt = NULL;
        
        if( acdbOpenAcDbEntity( pEnt, id, AcDb::kForWrite ) != Acad::eOk )
            continue;
            
        // on récupère la face
        AcDbFace* castEntity = static_cast<AcDbFace*>( pEnt );
        
        indexFaceI.resize( 0 );
        
        // on boucle sur les sommets
        for( int k = 0; k < 3; k++ )
        {
        
            // on récupère la position du sommet
            AcGePoint3d position;
            castEntity->getVertexAt( k, position );
            
            it = find( positions.begin(), positions.end(), position );
            
            if( it == positions.end() )
            {
                positions.push_back( position );
                indexFaceI.push_back( positions.size() - 1 );
            }
            
            else
                indexFaceI.push_back( it - positions.begin() );
        }
        
        index.push_back( indexFaceI );
        
        // on ferme la face
        castEntity->close();
    }
    
    // Liberer la selection
    acedSSFree( faceSelection );
    
    auto sizeP = positions.size();
    
    // on note les positions
    for( auto i = 0; i < sizeP; i++ )
    {
    
        ACHAR s1[255], s2[255], s3[255];
        acdbRToS( long double( positions[i][0] ), 2, 6, s1 );
        acdbRToS( long double( positions[i][1] ), 2, 6, s2 );
        acdbRToS( long double( positions[i][2] ), 2, 6, s3 );
        
        *file << "\n\t\t\t\t<P id=\"" + to_string( long long int( i ) ) +
            "\"> " + acStrToStr( s2 ) +
            " " + acStrToStr( s1 ) +
            " " + acStrToStr( s3 ) + "</P>";
    }
    
    *file << "\n\t\t\t</Pnts>";
    *file << "\n\t\t\t<Faces>";
    
    // on note les index
    int size = index.size();
    
    for( int i = 0; i < size; i++ )
    {
        *file << "\n\t\t\t\t<F n=\"" + to_string( long long int( i + sizeP ) ) + "\">" + to_string( long long int( index[i][0] ) ) + " " +
            to_string( long long int( index[i][1] ) ) + " " +
            to_string( long long int( index[i][2] ) ) + "</F>";
    }
    
    //Terminer le fichier
    *file << "\n\t\t\t</Faces>";
    *file << "\n\t\t</Definition>";
    *file << "\n\t</Surface>";
    
    //Sortir
    return length;
}


void exportFaceToXml(
    ofstream* file,
    const vector<FacesInLayer>& vecFL )
{
    //Recuperer la taille du vecteur
    long lenVecFL = vecFL.size();
    
    //Barre de progression
    ProgressBar prog = ProgressBar( _T( "Progression" ), lenVecFL );
    
    //Boucle sur la selection de faceinlayer
    for( int i = 0; i < lenVecFL; i++ )
    {
        //Progresser
        prog.moveUp( i );
        
        //Recuperer le ieme FL
        FacesInLayer fl = vecFL[i];
        
        //Recuperer la taille des faces dans le calques
        long lenFace = fl.facesIds.size();
        
        //Verifier si elle contient des faces
        if( lenFace == 0 )
            continue;
            
        //Completer le nom du calque
        *file << "\n\t<Surface name=\"" + acStrToStr( fl.layer ) + "\" desc=\"Description\">";
        *file << "\n\t\t<Definition surfType=\"TIN\">";
        *file << "\n\t\t\t<Pnts>";
        
        //Vecteur des positions de la selection
        vector<AcGePoint3d> positions;
        positions.resize( 0 );
        
        //Vecteur avec les index correspondants aux positions des sommets
        vector<vector<int>> index;
        
        //Vecteur avec les index de la face i
        vector<int> indexFaceI;
        index.resize( 0 );
        
        vector< AcGePoint3d >::iterator it;
        
        //Boucle sur les faces
        for( int f = 0; f < lenFace; f++ )
        {
            //Recuperer le feme face
            AcDbEntity* pEnt = NULL;
            
            if( acdbOpenAcDbEntity( pEnt, fl.facesIds[f], AcDb::kForRead ) != Acad::eOk )
                continue;
                
            //On recupere la face
            AcDbFace* face = static_cast<AcDbFace*>( pEnt );
            
            indexFaceI.resize( 0 );
            
            //On boucle sur les sommets
            for( int k = 0; k < 3; k++ )
            {
            
                // on récupère la position du sommet
                AcGePoint3d position;
                face->getVertexAt( k, position );
                
                it = find( positions.begin(), positions.end(), position );
                
                if( it == positions.end() )
                {
                    positions.push_back( position );
                    indexFaceI.push_back( positions.size() - 1 );
                }
                
                else
                    indexFaceI.push_back( it - positions.begin() );
            }
            
            index.push_back( indexFaceI );
            
            //On ferme la face
            face->close();
        }
        
        auto sizeP = positions.size();
        
        // on note les positions
        for( size_t i = 0; i < sizeP; i++ )
        {
        
            ACHAR s1[255], s2[255], s3[255];
            acdbRToS( long double( positions[i][0] ), 2, 6, s1 );
            acdbRToS( long double( positions[i][1] ), 2, 6, s2 );
            acdbRToS( long double( positions[i][2] ), 2, 6, s3 );
            
            *file << "\n\t\t\t\t<P id=\"" + to_string( long long int( i ) ) +
                "\"> " + acStrToStr( s2 ) +
                " " + acStrToStr( s1 ) +
                " " + acStrToStr( s3 ) + "</P>";
        }
        
        *file << "\n\t\t\t</Pnts>";
        *file << "\n\t\t\t<Faces>";
        
        // on note les index
        auto size = index.size();
        
        for( auto i = 0; i < size; i++ )
        {
            *file << "\n\t\t\t\t<F n=\"" + to_string( long long int( i + sizeP ) ) + "\">" + to_string( long long int( index[i][0] ) ) + " " +
                to_string( long long int( index[i][1] ) ) + " " +
                to_string( long long int( index[i][2] ) ) + "</F>";
        }
        
        //Terminer le fichier
        *file << "\n\t\t\t</Faces>";
        *file << "\n\t\t</Definition>";
        *file << "\n\t</Surface>";
    }
}

//Sous fonction pour la commande EXPORTLINE
void exportLines(
    AcDbLine*& _line,
    AcStringArray& _lData )
{
    //--------------------------------------------<Variables>
    //Type de calque
    AcString line_LayerType,
             //Type de ligne
             line_LineType,
             //Matériel
             line_MaterialName,
             //Taille de ligne
             line_LineWeightStr,
             //Couleur
             line_Color,
             //Handle
             line_Handle,
             //Transparence
             line_TranspStr,
             
             //Nom du style de tracé
             line_PlotStyleName;
             
    //Taille du ligne de la line en int
    int line_LineWeigth;
    
    //Transparence de la line
    AcCmTransparency line_Transp;
    
    //Longueur, linetype scale
    double line_Length,
           line_LineTypeScale,
           line_Thickness;
    //--------------------------------------------</Variables>
    
    
    //--------------------------------------------<Traitement>
    
    //Prendre les paramètres de l'entité
    getEntityParams(
        _line,
        line_LayerType,
        line_LineType,
        line_MaterialName,
        line_Color,
        line_Handle,
        line_LineWeigth,
        line_LineWeightStr,
        line_Transp,
        line_TranspStr,
        line_LineTypeScale,
        line_PlotStyleName );
        
    //Prendre la lonngueur de la polyligne
    line_Length = getLength( _line );
    
    //Thickness
    line_Thickness = _line->thickness();
    
    string thickness = to_string( line_Thickness );
    string lineTypeScale = to_string( line_LineTypeScale );
    string lineLength = to_string( line_Length );
    
    //Coordonnées du point de départ
    string dep_x = to_string( _line->startPoint().x );
    string dep_y = to_string( _line->startPoint().y );
    string dep_z = to_string( _line->startPoint().z );
    //Coordonnées du point d'arrivée
    string end_x = to_string( _line->endPoint().x );
    string end_y = to_string( _line->endPoint().y );
    string end_z = to_string( _line->endPoint().z );
    
    
    //Thickness de la polyligne 2d
    line_Thickness = _line->thickness();
    
    _lData.push_back( line_Handle );
    _lData.push_back( line_LayerType );
    _lData.push_back( line_Color );
    _lData.push_back( line_LineType );
    _lData.push_back( strToAcStr( lineTypeScale ) );
    _lData.push_back( line_LineWeightStr );
    _lData.push_back( line_TranspStr );
    _lData.push_back( strToAcStr( thickness ) );
    _lData.push_back( line_MaterialName );
    _lData.push_back( strToAcStr( lineLength ) );
    _lData.push_back( strToAcStr( dep_x ) );
    _lData.push_back( strToAcStr( dep_y ) );
    _lData.push_back( strToAcStr( dep_z ) );
    _lData.push_back( strToAcStr( end_x ) );
    _lData.push_back( strToAcStr( end_y ) );
    _lData.push_back( strToAcStr( end_z ) );
    
    
    //--------------------------------------------</Traitement>
}


void exportCircles(
    AcDbCircle*& _circle,
    AcStringArray& _cData )
{
    //--------------------------------------------<Variables>
    //Type de calque
    AcString circle_LayerType,
             //Type de ligne
             circle_LineType,
             //Matériel
             circle_MaterialName,
             //Taille de ligne
             circle_LineWeightStr,
             //Couleur
             circle_Color,
             //Handle
             circle_Handle,
             //Transparence
             circle_TranspStr,
             
             //Nom du style de tracé
             circle_PlotStyleName;
             
    //Taille du ligne du cercle en int
    int circle_LineWeigth;
    
    //Transparence du cercle
    AcCmTransparency circle_Transp;
    
    //Centre du cercle
    AcGePoint3d circle_center = AcGePoint3d::kOrigin;
    
    //Longueur, linetype scale
    double circle_Radius,
           circle_LineTypeScale,
           circle_Thickness;
    //--------------------------------------------</Variables>
    
    
    //--------------------------------------------<Traitement>
    
    //Prendre les paramètres de l'entité
    getEntityParams(
        _circle,
        circle_LayerType,
        circle_LineType,
        circle_MaterialName,
        circle_Color,
        circle_Handle,
        circle_LineWeigth,
        circle_LineWeightStr,
        circle_Transp,
        circle_TranspStr,
        circle_LineTypeScale,
        circle_PlotStyleName );
        
    //Determiner la position du centre
    circle_center = _circle->center();
    
    //Determiner le diamètre du cercle
    circle_Radius = _circle->radius();
    
    //Position du centre du cercle
    string centre_x = to_string( circle_center.x );
    string centre_y = to_string( circle_center.y );
    string centre_z = to_string( circle_center.z );
    string rayon = to_string( circle_Radius );
    
    //La surface du cercle
    double c_area ;
    _circle->getArea( c_area );
    
    string surface = to_string( c_area );
    
    
    //Thickness du cercle
    circle_Thickness = _circle->thickness();
    string thickness = to_string( circle_Thickness );
    string linetypescale = to_string( circle_LineTypeScale );
    
    //Ecrire les données dans le tableau
    _cData.push_back( circle_Handle );
    _cData.push_back( circle_LayerType );
    _cData.push_back( circle_Color );
    _cData.push_back( circle_LineType );
    _cData.push_back( strToAcStr( linetypescale ) );
    _cData.push_back( circle_LineWeightStr );
    _cData.push_back( circle_TranspStr );
    _cData.push_back( strToAcStr( thickness ) );
    _cData.push_back( circle_MaterialName );
    _cData.push_back( strToAcStr( centre_x ) );
    _cData.push_back( strToAcStr( centre_y ) );
    _cData.push_back( strToAcStr( centre_z ) );
    _cData.push_back( strToAcStr( rayon ) );
    _cData.push_back( strToAcStr( surface ) );
    
    
    //--------------------------------------------</Traitement>
}

void exportHatch(
    AcDbHatch*& _hatch,
    AcStringArray& _hData )
{
    //--------------------------------------------<Variables>
    //Type de calque
    AcString hatch_LayerType,
             //Type de ligne
             hatch_LineType,
             //Matériel
             hatch_MaterialName,
             //Taille de ligne
             hatch_LineWeightStr,
             //Couleur
             hatch_Color,
             //Handle
             hatch_Handle,
             //Transparence
             hatch_TranspStr,
             
             //Nom du style de tracé
             hatch_PlotStyleName;
             
    //Taille du ligne du cercle en int
    int hatch_LineWeigth;
    
    //Transparence du cercle
    AcCmTransparency hatch_Transp;
    
    //Centre du cercle
    AcGePoint3d hatch_center = AcGePoint3d::kOrigin;
    
    //Longueur, linetype scale
    double hatch_Radius,
           hatch_LineTypeScale;
    //--------------------------------------------</Variables>
    
    
    //--------------------------------------------<Traitement>
    
    //Prendre les paramètres de l'entité
    getEntityParams(
        _hatch,
        hatch_LayerType,
        hatch_LineType,
        hatch_MaterialName,
        hatch_Color,
        hatch_Handle,
        hatch_LineWeigth,
        hatch_LineWeightStr,
        hatch_Transp,
        hatch_TranspStr,
        hatch_LineTypeScale,
        hatch_PlotStyleName );
        
    //Recuperer les coordonnées du point d'insertion des hachures
    AcDbExtents b;
    _hatch->bounds( b );
    AcGePoint3d pos = midPoint3d( b.minPoint(), b.maxPoint() );
    
    
    //Recuperer l'angle d'inclinaison
    double angle = _hatch->patternAngle();
    
    //Recuperer la surface
    double area;
    _hatch->getArea( area );
    
    //Recuperer le motif
    AcString pattern = _hatch->patternName();
    
    string linetypescale = to_string( hatch_LineTypeScale );
    
    _hData.push_back( hatch_Handle );
    _hData.push_back( hatch_LayerType );
    _hData.push_back( hatch_Color );
    _hData.push_back( hatch_LineType );
    _hData.push_back( strToAcStr( linetypescale ) );
    _hData.push_back( hatch_LineWeightStr );
    _hData.push_back( hatch_TranspStr );
    _hData.push_back( hatch_MaterialName );
    _hData.push_back( strToAcStr( to_string( pos.x ) ) );
    _hData.push_back( strToAcStr( to_string( pos.y ) ) );
    _hData.push_back( strToAcStr( to_string( pos.z ) ) );
    _hData.push_back( strToAcStr( to_string( angle ) ) );
    _hData.push_back( strToAcStr( to_string( area ) ) );
    _hData.push_back( pattern );
    
    
    
    //--------------------------------------------</Traitement>
}

/* ==> LINE SHP <== */


vector<AcString> createLineField( const AcString& paramFile )
{

    //Listes des champs pour les blocks
    vector<AcString> field;
    vector<AcString> lastField;
    field.push_back( _T( "HANDLE" ) );
    field.push_back( _T( "LAYER" ) );
    field.push_back( _T( "COLOR" ) );
    field.push_back( _T( "LINETYPE" ) );
    field.push_back( _T( "LTSCALE" ) );
    field.push_back( _T( "LWEIGHT" ) );
    field.push_back( _T( "TRANSPARENCY" ) );
    field.push_back( _T( "THICKNESS" ) );
    field.push_back( _T( "MATERIAL" ) );
    field.push_back( _T( "LENGTH2D" ) );
    field.push_back( _T( "STARTX" ) );
    field.push_back( _T( "STARTY" ) );
    field.push_back( _T( "STARTZ" ) );
    field.push_back( _T( "ENDX" ) );
    field.push_back( _T( "ENDY" ) );
    field.push_back( _T( "ENDZ" ) );
    
    //Charcger le fichier de configuration
    Book* book = xlCreateBook();
    
    vector<AcString> fildOption[2];
    
    if( book->load( paramFile ) )
    {
        Sheet* sheet = book->getSheet( 0 );
        
        if( sheet )
            for( int row = sheet->firstRow(); row < sheet->lastRow(); ++row )
                for( int col = sheet->firstCol(); col < sheet->lastCol(); col++ )
                    fildOption[col].push_back( sheet->readStr( row, col ) );
    }
    
    else
    {
        //Assurer que la longueur du titre ne depasse pas de 10
        for( int i = 0; i < field.size(); i++ )
        {
            string tempField = acStrToStr( field[i] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        return lastField;
    }
    
    vector<AcString>::iterator it = fildOption[1].begin();
    
    for( it; it != fildOption[1].end(); ++it )
    {
        //Recuperer le numero de la ligne
        int index = it - fildOption[1].begin();
        string tempField = acStrToStr( *it );
        
        if( ( it->compareNoCase( "" ) != 0 ) && ( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() ) )
        {
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        else if( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() )
        {
            tempField = acStrToStr( fildOption[0][index] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
    }
    
    //Fermer le fichier excel
    book->release();
    return lastField;
}

vector<AcString> createPolyField( const AcString& paramFile )
{

    //Listes des champs pour les blocks
    vector<AcString> field;
    vector<AcString> lastField;
    field.push_back( _T( "HANDLE" ) );
    field.push_back( _T( "LAYER" ) );
    field.push_back( _T( "COLOR" ) );
    field.push_back( _T( "LINETYPE" ) );
    field.push_back( _T( "LTSCALE" ) );
    field.push_back( _T( "LWEIGHT" ) );
    field.push_back( _T( "TRANSPARENCY" ) );
    field.push_back( _T( "THICKNESS" ) );
    field.push_back( _T( "MATERIAL" ) );
    field.push_back( _T( "LENGTH2D" ) );
    field.push_back( _T( "LENGTH3D" ) );
    field.push_back( _T( "CLOSED" ) );
    
    
    //Charcger le fichier de configuration
    Book* book = xlCreateBook();
    
    vector<AcString> fildOption[2];
    
    if( book->load( paramFile ) )
    {
        Sheet* sheet = book->getSheet( 0 );
        
        if( sheet )
            for( int row = sheet->firstRow(); row < sheet->lastRow(); ++row )
                for( int col = sheet->firstCol(); col < sheet->lastCol(); col++ )
                    fildOption[col].push_back( sheet->readStr( row, col ) );
    }
    
    else
    {
        //Assurer que la longueur du titre ne depasse pas de 10
        for( int i = 0; i < field.size(); i++ )
        {
            string tempField = acStrToStr( field[i] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        return lastField;
    }
    
    vector<AcString>::iterator it = fildOption[1].begin();
    
    for( it; it != fildOption[1].end(); ++it )
    {
        //Recuperer le numero de la ligne
        int index = it - fildOption[1].begin();
        string tempField = acStrToStr( *it );
        
        if( ( it->compareNoCase( "" ) != 0 ) && ( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() ) )
        {
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        else if( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() )
        {
            tempField = acStrToStr( fildOption[0][index] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
    }
    
    //Fermer le fichier excel
    book->release();
    return lastField;
}

vector<AcString> createClosedPolyField( const AcString& paramFile )
{

    //Listes des champs pour les blocks
    vector<AcString> field;
    vector<AcString> lastField;
    field.push_back( _T( "HANDLE" ) );
    field.push_back( _T( "LAYER" ) );
    field.push_back( _T( "COLOR" ) );
    field.push_back( _T( "LINETYPE" ) );
    field.push_back( _T( "LTSCALE" ) );
    field.push_back( _T( "LWEIGHT" ) );
    field.push_back( _T( "TRANSPARENCY" ) );
    field.push_back( _T( "THICKNESS" ) );
    field.push_back( _T( "MATERIAL" ) );
    field.push_back( _T( "LENGTH2D" ) );
    field.push_back( _T( "LENGTH3D" ) );
    field.push_back( _T( "AREA" ) );
    
    
    //Charcger le fichier de configuration
    Book* book = xlCreateBook();
    
    vector<AcString> fildOption[2];
    
    if( book->load( paramFile ) )
    {
        Sheet* sheet = book->getSheet( 0 );
        
        if( sheet )
            for( int row = sheet->firstRow(); row < sheet->lastRow(); ++row )
                for( int col = sheet->firstCol(); col < sheet->lastCol(); col++ )
                    fildOption[col].push_back( sheet->readStr( row, col ) );
    }
    
    else
    {
        //Assurer que la longueur du titre ne depasse pas de 10
        for( int i = 0; i < field.size(); i++ )
        {
            string tempField = acStrToStr( field[i] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        return lastField;
    }
    
    vector<AcString>::iterator it = fildOption[1].begin();
    
    for( it; it != fildOption[1].end(); ++it )
    {
        //Recuperer le numero de la ligne
        int index = it - fildOption[1].begin();
        string tempField = acStrToStr( *it );
        
        if( ( it->compareNoCase( "" ) != 0 ) && ( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() ) )
        {
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        else if( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() )
        {
            tempField = acStrToStr( fildOption[0][index] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
    }
    
    //Fermer le fichier excel
    book->release();
    return lastField;
}

vector<AcString> createCircleField( const AcString& paramFile )
{

    //Listes des champs pour les blocks
    vector<AcString> field;
    vector<AcString> lastField;
    field.push_back( _T( "HANDLE" ) );
    field.push_back( _T( "LAYER" ) );
    field.push_back( _T( "COLOR" ) );
    field.push_back( _T( "LINETYPE" ) );
    field.push_back( _T( "LTSCALE" ) );
    field.push_back( _T( "LWEIGHT" ) );
    field.push_back( _T( "TRANSPARENCY" ) );
    field.push_back( _T( "THICKNESS" ) );
    field.push_back( _T( "MATERIAL" ) );
    field.push_back( _T( "POSX" ) );
    field.push_back( _T( "POSY" ) );
    field.push_back( _T( "POSZ" ) );
    field.push_back( _T( "RADIUS" ) );
    
    //Charcger le fichier de configuration
    Book* book = xlCreateBook();
    
    vector<AcString> fildOption[2];
    
    if( book->load( paramFile ) )
    {
        Sheet* sheet = book->getSheet( 0 );
        
        if( sheet )
            for( int row = sheet->firstRow(); row < sheet->lastRow(); ++row )
                for( int col = sheet->firstCol(); col < sheet->lastCol(); col++ )
                    fildOption[col].push_back( sheet->readStr( row, col ) );
    }
    
    else
    {
        //Assurer que la longueur du titre ne depasse pas de 10
        for( int i = 0; i < field.size(); i++ )
        {
            string tempField = acStrToStr( field[i] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        return lastField;
    }
    
    vector<AcString>::iterator it = fildOption[1].begin();
    
    for( it; it != fildOption[1].end(); ++it )
    {
        //Recuperer le numero de la ligne
        int index = it - fildOption[1].begin();
        string tempField = acStrToStr( *it );
        
        if( ( it->compareNoCase( "" ) != 0 ) && ( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() ) )
        {
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        else if( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() )
        {
            tempField = acStrToStr( fildOption[0][index] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
    }
    
    //Fermer le fichier excel
    book->release();
    return lastField;
}


vector<AcString> createPoly3dField( const AcString& paramFile )
{

    //Listes des champs pour les blocks
    vector<AcString> field;
    vector<AcString> lastField;
    field.push_back( _T( "HANDLE" ) );
    field.push_back( _T( "LAYER" ) );
    field.push_back( _T( "COLOR" ) );
    field.push_back( _T( "LINETYPE" ) );
    field.push_back( _T( "LTSCALE" ) );
    field.push_back( _T( "LWEIGHT" ) );
    field.push_back( _T( "TRANSPARENCY" ) );
    field.push_back( _T( "MATERIAL" ) );
    field.push_back( _T( "LENGTH2D" ) );
    field.push_back( _T( "LENGTH3D" ) );
    field.push_back( _T( "CLOSED" ) );
    
    
    //Charcger le fichier de configuration
    Book* book = xlCreateBook();
    
    vector<AcString> fildOption[2];
    
    if( book->load( paramFile ) )
    {
        Sheet* sheet = book->getSheet( 0 );
        
        if( sheet )
            for( int row = sheet->firstRow(); row < sheet->lastRow(); ++row )
                for( int col = sheet->firstCol(); col < sheet->lastCol(); col++ )
                    fildOption[col].push_back( sheet->readStr( row, col ) );
    }
    
    else
    {
        //Assurer que la longueur du titre ne depasse pas de 10
        for( int i = 0; i < field.size(); i++ )
        {
            string tempField = acStrToStr( field[i] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        return lastField;
    }
    
    vector<AcString>::iterator it = fildOption[1].begin();
    
    for( it; it != fildOption[1].end(); ++it )
    {
        //Recuperer le numero de la ligne
        int index = it - fildOption[1].begin();
        string tempField = acStrToStr( *it );
        
        if( ( it->compareNoCase( "" ) != 0 ) && ( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() ) )
        {
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        else if( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() )
        {
            tempField = acStrToStr( fildOption[0][index] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
    }
    
    //Fermer le fichier excel
    book->release();
    return lastField;
}

vector<AcString> createClosedPoly3dField( const AcString& paramFile )
{

    //Listes des champs pour les blocks
    vector<AcString> field;
    vector<AcString> lastField;
    field.push_back( _T( "HANDLE" ) );
    field.push_back( _T( "LAYER" ) );
    field.push_back( _T( "COLOR" ) );
    field.push_back( _T( "LINETYPE" ) );
    field.push_back( _T( "LTSCALE" ) );
    field.push_back( _T( "LWEIGHT" ) );
    field.push_back( _T( "TRANSPARENCY" ) );
    field.push_back( _T( "MATERIAL" ) );
    field.push_back( _T( "LENGTH2D" ) );
    field.push_back( _T( "LENGTH3D" ) );
    field.push_back( _T( "AREA" ) );
    
    
    //Charcger le fichier de configuration
    Book* book = xlCreateBook();
    
    vector<AcString> fildOption[2];
    
    if( book->load( paramFile ) )
    {
        Sheet* sheet = book->getSheet( 0 );
        
        if( sheet )
            for( int row = sheet->firstRow(); row < sheet->lastRow(); ++row )
                for( int col = sheet->firstCol(); col < sheet->lastCol(); col++ )
                    fildOption[col].push_back( sheet->readStr( row, col ) );
    }
    
    else
    {
        //Assurer que la longueur du titre ne depasse pas de 10
        for( int i = 0; i < field.size(); i++ )
        {
            string tempField = acStrToStr( field[i] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        return lastField;
    }
    
    vector<AcString>::iterator it = fildOption[1].begin();
    
    for( it; it != fildOption[1].end(); ++it )
    {
        //Recuperer le numero de la ligne
        int index = it - fildOption[1].begin();
        string tempField = acStrToStr( *it );
        
        if( ( it->compareNoCase( "" ) != 0 ) && ( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() ) )
        {
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        else if( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() )
        {
            tempField = acStrToStr( fildOption[0][index] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
    }
    
    //Fermer le fichier excel
    book->release();
    return lastField;
}

void drawLineShp( AcDbLine* line,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field )
{
    double* x;
    double* y;
    double* z;
    
    x = new double[2];
    y = new double[2];
    z = new double[2];
    
    x[0] = line->startPoint().x;
    y[0] = line->startPoint().y;
    z[0] = line->startPoint().z;
    
    x[1] = line->endPoint().x;
    y[1] = line->endPoint().y;
    z[1] = line->endPoint().z;
    
    
    if( ( x != NULL ) && ( y != NULL ) && ( z != NULL ) )
    {
        //
        SHPObject* object = SHPCreateObject( SHPT_ARCZ, -1, 1, NULL, NULL, 2, x, y, z, NULL );
        int iShape = SHPWriteObject( myShapeFile, -1, object );
        
        //Fermer la shapefile
        SHPDestroyObject( object );
        
        //Ecrire le fichier dfg
        writeLine( line, dbfHandle, iShape, field );
    }
}

void drawLineShp( AcDbPolyline* poly2d,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field )
{
    double* x;
    double* y;
    double* z;
    
    int vertNumber = poly2d->numVerts();
    
    if( poly2d->isClosed() )
        vertNumber++;
        
    x = new double[vertNumber];
    y = new double[vertNumber];
    z = new double[vertNumber];
    
    //Boucler sur les sommets
    for( int j = 0; j < poly2d->numVerts(); j++ )
    {
        AcGePoint3d pt;
        poly2d->getPointAt( j, pt );
        x[j] = pt.x;
        y[j] = pt.y;
        z[j] = pt.z;
    }
    
    if( poly2d->isClosed() )
    {
        AcGePoint3d pt;
        poly2d->getPointAt( 0, pt );
        x[vertNumber - 1] = pt.x;
        y[vertNumber - 1] = pt.y;
        z[vertNumber - 1] = pt.z;
    }
    
    if( ( x != NULL ) && ( y != NULL ) && ( z != NULL ) )
    {
        SHPObject* object = SHPCreateObject( SHPT_ARCZ, -1, 1, NULL, NULL, vertNumber, x, y, z, NULL );
        int iShape = SHPWriteObject( myShapeFile, -1, object );
        //Fermer la shapefile
        SHPDestroyObject( object );
        //Ecrire le fichier dfg
        writePoly2d( poly2d, dbfHandle, iShape, field );
    }
}

void drawCircleShp( AcDbCircle* circle,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field )
{
    double* x;
    double* y;
    double* z;
    
    //Recuperer tous les points du cercle
    AcGePoint3dArray arrPoint;
    
    double fleche;
    
    arrPoint = discretize( circle, 0.00001 );
    
    //Recuperer le premier element du tableau
    AcGePoint3d pt = arrPoint[0];
    
    //Ajouter en fin du tableau le contenu du premier indice du tableau
    arrPoint.append( pt );
    
    //Recuperer la taille du tableau
    int vertNumber = arrPoint.size();
    
    x = new double[vertNumber];
    y = new double[vertNumber];
    z = new double[vertNumber];
    
    //Boucler sur les sommets
    for( int j = 0; j < vertNumber; j++ )
    {
        AcGePoint3d pt = arrPoint[j];
        x[j] = pt.x;
        y[j] = pt.y;
        z[j] = pt.z;
    }
    
    if( ( x != NULL ) && ( y != NULL ) && ( z != NULL ) )
    {
        SHPObject* object = SHPCreateObject( SHPT_ARCZ, -1, 1, NULL, NULL, vertNumber, x, y, z, NULL );
        int iShape = SHPWriteObject( myShapeFile, -1, object );
        //Fermer la shapefile
        SHPDestroyObject( object );
        //Ecrire le fichier dfg
        writeCircle( circle, dbfHandle, iShape, field );
    }
}


void drawClosedPolyShp(
    AcDbPolyline* poly2d,
    vector<AcGePoint3d>& pt,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field )
{


    int vertNumber = pt.size();
    
    double* pointX = new double[vertNumber];
    double* pointY = new double[vertNumber];
    double* pointZ = new double[vertNumber];
    
    
    
    //Boucler sur la listes des points
    for( int i = 0; i < pt.size(); i++ )
    {
        pointX[i] = pt[i].x;
        pointY[i] = pt[i].y;
        pointZ[i] = pt[i].z;
        
    }
    
    if( ( pointX != NULL ) && ( pointY != NULL ) && ( pointZ != NULL ) )
    {
        SHPObject* object = SHPCreateObject( SHPT_POLYGONZ, -1, 1, NULL, NULL, vertNumber, pointX, pointY, pointZ, NULL );
        int iShape = SHPWriteObject( myShapeFile, -1, object );
        
        //Fermer la shapefile
        SHPDestroyObject( object );
        
        //Ecrire le fichier dfg
        writeClosedPoly2d( poly2d, dbfHandle, iShape, field );
    }
}




void drawClosedPolyShp( AcDb3dPolyline* poly3d,
    vector<AcGePoint3d>& pt,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field )
{
    double* pointX;
    double* pointY;
    double* pointZ;
    
    int vertNumber = pt.size();
    
    pointX = new double[vertNumber];
    pointY = new double[vertNumber];
    pointZ = new double[vertNumber];
    
    SHPObject* object;
    
    int i = 0;
    
    //Boucler sur la listes des points
    for( AcGePoint3d p : pt )
    {
        pointX[i] = p.x;
        pointY[i] = p.y;
        pointZ[i] = p.z;
        i++;
    }
    
    if( ( pointX != NULL ) && ( pointY != NULL ) && ( pointZ != NULL ) )
    {
        SHPObject* object = SHPCreateObject( SHPT_POLYGONZ, -1, 1, NULL, NULL, vertNumber, pointX, pointY, pointZ, NULL );
        int iShape = SHPWriteObject( myShapeFile, -1, object );
        
        //Fermer la shapefile
        SHPDestroyObject( object );
        
        //Ecrire le fichier dfg
        writeClosedPoly3d( poly3d, dbfHandle, iShape, field );
    }
}



void drawLineShp( AcDb3dPolyline* poly3d,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field )
{
    double* x;
    double* y;
    double* z;
    
    int vertNumber = getNumberOfVertex( poly3d );
    
    AcDbObjectIterator* iterPoly = poly3d->vertexIterator();
    AcDb3dPolylineVertex* vertex;
    int j = 0;
    
    if( poly3d->isClosed() )
        vertNumber++;
        
    x = new double[vertNumber];
    y = new double[vertNumber];
    z = new double[vertNumber];
    
    
    for( iterPoly->start(); !iterPoly->done(); iterPoly->step() )
    {
        if( Acad::eOk == poly3d->openVertex( vertex, iterPoly->objectId(), AcDb::kForRead ) )
        {
            x[j] = vertex->position().x;
            y[j] = vertex->position().y;
            z[j] = vertex->position().z;
        }
        
        j++;
    }
    
    vertex->close();
    
    if( poly3d->isClosed() )
    {
        AcGePoint3d p;
        poly3d->getStartPoint( p );
        x[vertNumber - 1] = p.x;
        y[vertNumber - 1] = p.y;
        z[vertNumber - 1] = p.z;
    }
    
    
    if( ( x != NULL ) && ( y != NULL ) && ( z != NULL ) )
    {
        SHPObject* object = SHPCreateObject( SHPT_ARCZ, -1, 1, NULL, NULL, vertNumber, x, y, z, NULL );
        int iShape = SHPWriteObject( myShapeFile, -1, object );
        
        //Fermer la shapefile
        SHPDestroyObject( object );
        
        //Ecrire le fichier dfg
        writePoly3d( poly3d, dbfHandle, iShape, field );
    }
}


/* ==> END LINE SHP <== */


/* ==> TEXT SHAPE <== */

vector<AcString> createTextField( const AcString& paramFile )
{
    //Listes des champs pour les blocks
    vector<AcString> field;
    vector<AcString> lastField;
    
    field.push_back( _T( "HANDLE" ) );
    field.push_back( _T( "LAYER" ) );
    field.push_back( _T( "COLOR" ) );
    field.push_back( _T( "LINETYPE" ) );
    field.push_back( _T( "LTSCALE" ) );
    field.push_back( _T( "LWEIGHT" ) );
    field.push_back( _T( "TRANSPARENCY" ) );
    field.push_back( _T( "MATERIAL" ) );
    field.push_back( _T( "POSX" ) );
    field.push_back( _T( "POSY" ) );
    field.push_back( _T( "POSZ" ) );
    field.push_back( _T( "ANGLE" ) );
    field.push_back( _T( "VALUE" ) );
    field.push_back( _T( "HEIGHT" ) );
    field.push_back( _T( "STYLE" ) );
    field.push_back( _T( "ALIGN" ) );
    
    //Charcger le fichier de configuration
    Book* book = xlCreateBook();
    
    vector<AcString> fildOption[2];
    
    if( book->load( paramFile ) )
    {
        Sheet* sheet = book->getSheet( 0 );
        
        if( sheet )
            for( int row = sheet->firstRow(); row < sheet->lastRow(); ++row )
                for( int col = sheet->firstCol(); col < sheet->lastCol(); col++ )
                    fildOption[col].push_back( sheet->readStr( row, col ) );
    }
    
    else
    {
    
        for( int i = 0; i < field.size(); i++ )
        {
            string tempField = acStrToStr( field[i] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        return lastField;
    }
    
    vector<AcString>::iterator it = fildOption[1].begin();
    
    for( it; it != fildOption[1].end(); ++it )
    {
        //Recuperer le numero de la ligne
        int index = it - fildOption[1].begin();
        string tempField = acStrToStr( *it );
        
        if( ( it->compareNoCase( "" ) != 0 ) && ( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() ) )
        {
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        else if( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() )
        {
            tempField = acStrToStr( fildOption[0][index] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
    }
    
    //Fermer le fichier excel
    book->release();
    return lastField;
}

vector<AcString> createTextPoly( const AcString& paramFile )
{
    //Listes des champs pour les blocks
    vector<AcString> field;
    vector<AcString> lastField;
    
    field.push_back( _T( "HANDLE" ) );
    field.push_back( _T( "LAYER" ) );
    field.push_back( _T( "COLOR" ) );
    field.push_back( _T( "LINETYPE" ) );
    field.push_back( _T( "LTSCALE" ) );
    field.push_back( _T( "LWEIGHT" ) );
    field.push_back( _T( "TRANSPARENCY" ) );
    field.push_back( _T( "MATERIAL" ) );
    field.push_back( _T( "X" ) );
    field.push_back( _T( "Y" ) );
    field.push_back( _T( "Z" ) );
    field.push_back( _T( "ANGLE" ) );
    field.push_back( _T( "VALUE" ) );
    field.push_back( _T( "HEIGHT" ) );
    field.push_back( _T( "STYLE" ) );
    field.push_back( _T( "ALIGN" ) );
    field.push_back( _T( "POLYHANDLE" ) );
    field.push_back( _T( "PLENGTH2D" ) );
    field.push_back( _T( "POLYAREA" ) );
    
    
    //Charcger le fichier de configuration
    Book* book = xlCreateBook();
    
    vector<AcString> fildOption[2];
    
    if( book->load( paramFile ) )
    {
        Sheet* sheet = book->getSheet( 0 );
        
        if( sheet )
            for( int row = sheet->firstRow(); row < sheet->lastRow(); ++row )
                for( int col = sheet->firstCol(); col < sheet->lastCol(); col++ )
                    fildOption[col].push_back( sheet->readStr( row, col ) );
    }
    
    else
    {
    
        for( int i = 0; i < field.size(); i++ )
        {
            string tempField = acStrToStr( field[i] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        return lastField;
    }
    
    vector<AcString>::iterator it = fildOption[1].begin();
    
    for( it; it != fildOption[1].end(); ++it )
    {
        //Recuperer le numero de la ligne
        int index = it - fildOption[1].begin();
        string tempField = acStrToStr( *it );
        
        if( ( it->compareNoCase( "" ) != 0 ) && ( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() ) )
        {
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        else if( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() )
        {
            tempField = acStrToStr( fildOption[0][index] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
    }
    
    //Fermer le fichier excel
    book->release();
    return lastField;
}

vector<AcString> createTextPoly3d( const AcString& paramFile )
{
    //Listes des champs pour les blocks
    vector<AcString> field;
    vector<AcString> lastField;
    
    field.push_back( _T( "HANDLE" ) );
    field.push_back( _T( "LAYER" ) );
    field.push_back( _T( "COLOR" ) );
    field.push_back( _T( "LINETYPE" ) );
    field.push_back( _T( "LTSCALE" ) );
    field.push_back( _T( "LWEIGHT" ) );
    field.push_back( _T( "TRANSPARENCY" ) );
    field.push_back( _T( "MATERIAL" ) );
    field.push_back( _T( "X" ) );
    field.push_back( _T( "Y" ) );
    field.push_back( _T( "Z" ) );
    field.push_back( _T( "ANGLE" ) );
    field.push_back( _T( "VALUE" ) );
    field.push_back( _T( "HEIGHT" ) );
    field.push_back( _T( "STYLE" ) );
    field.push_back( _T( "ALIGN" ) );
    field.push_back( _T( "POLYHANDLE" ) );
    field.push_back( _T( "PLENGTH2D" ) );
    field.push_back( _T( "PLENGTH3D" ) );
    field.push_back( _T( "POLYAREA" ) );
    
    
    //Charcger le fichier de configuration
    Book* book = xlCreateBook();
    
    vector<AcString> fildOption[2];
    
    if( book->load( paramFile ) )
    {
        Sheet* sheet = book->getSheet( 0 );
        
        if( sheet )
            for( int row = sheet->firstRow(); row < sheet->lastRow(); ++row )
                for( int col = sheet->firstCol(); col < sheet->lastCol(); col++ )
                    fildOption[col].push_back( sheet->readStr( row, col ) );
    }
    
    else
    {
    
        for( int i = 0; i < field.size(); i++ )
        {
            string tempField = acStrToStr( field[i] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        return lastField;
    }
    
    vector<AcString>::iterator it = fildOption[1].begin();
    
    for( it; it != fildOption[1].end(); ++it )
    {
        //Recuperer le numero de la ligne
        int index = it - fildOption[1].begin();
        string tempField = acStrToStr( *it );
        
        if( ( it->compareNoCase( "" ) != 0 ) && ( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() ) )
        {
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        else if( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() )
        {
            tempField = acStrToStr( fildOption[0][index] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
    }
    
    //Fermer le fichier excel
    book->release();
    return lastField;
}

void drawTextShp( AcDbText* txt,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field )
{
    double* pointX;
    double* pointY;
    double* pointZ;
    
    SHPObject* object;
    
    //Recuperer le point d'insertino du texte
    AcGePoint3d txtInsertPoint = txt->position();
    
    pointX = &txtInsertPoint.x;
    pointY = &txtInsertPoint.y;
    pointZ = &txtInsertPoint.z;
    
    object = SHPCreateObject( SHPT_POINTZ, -1, 1, NULL, NULL, 1, pointX, pointY, pointZ, NULL );
    int iShape = SHPWriteObject( myShapeFile, -1, object );
    
    //Ecrire le dwg
    writeText( txt, dbfHandle, iShape, field );
    
    //Fermer la shapefile
    SHPDestroyObject( object );
    
    //Fermer le texte
    txt->close();
}


void drawMTextShp( AcDbMText* txt,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field )
{
    double* pointX;
    double* pointY;
    double* pointZ;
    
    SHPObject* object;
    
    //Recuperer le point d'insertino du texte
    AcGePoint3d txtInsertPoint = txt->location();
    
    pointX = &txtInsertPoint.x;
    pointY = &txtInsertPoint.y;
    pointZ = &txtInsertPoint.z;
    
    object = SHPCreateObject( SHPT_POINTZ, -1, 1, NULL, NULL, 1, pointX, pointY, pointZ, NULL );
    int iShape = SHPWriteObject( myShapeFile, -1, object );
    
    //Ecrire le dwg
    writeText( txt, dbfHandle, iShape, field );
    
    //Fermer la shapefile
    SHPDestroyObject( object );
    
    //Fermer le texte
    txt->close();
}


/* ==> END TEXT SHAPE <==  */

/* ==> POINT SHAPE <== */

vector<AcString> createPointField( const AcString& paramFile )
{

    //Listes des champs pour les blocks
    vector<AcString> field;
    vector<AcString> lastField;
    
    field.push_back( _T( "HANDLE" ) );
    field.push_back( _T( "LAYER" ) );
    field.push_back( _T( "COLOR" ) );
    field.push_back( _T( "LINETYPE" ) );
    field.push_back( _T( "LTSCALE" ) );
    field.push_back( _T( "LWEIGHT" ) );
    field.push_back( _T( "TRANSPARENCY" ) );
    field.push_back( _T( "THICKNESS" ) );
    field.push_back( _T( "MATERIAL" ) );
    field.push_back( _T( "POSX" ) );
    field.push_back( _T( "POSY" ) );
    field.push_back( _T( "POSZ" ) );
    
    
    //Charcger le fichier de configuration
    Book* book = xlCreateBook();
    
    vector<AcString> fildOption[2];
    
    if( book->load( paramFile ) )
    {
        Sheet* sheet = book->getSheet( 0 );
        
        if( sheet )
            for( int row = sheet->firstRow(); row < sheet->lastRow(); ++row )
                for( int col = sheet->firstCol(); col < sheet->lastCol(); col++ )
                    fildOption[col].push_back( sheet->readStr( row, col ) );
    }
    
    else
    {
    
        for( int i = 0; i < field.size(); i++ )
        {
            string tempField = acStrToStr( field[i] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        return lastField;
    }
    
    vector<AcString>::iterator it = fildOption[1].begin();
    
    for( it; it != fildOption[1].end(); ++it )
    {
        //Recuperer le numero de la ligne
        int index = it - fildOption[1].begin();
        string tempField = acStrToStr( *it );
        
        if( ( it->compareNoCase( "" ) != 0 ) && ( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() ) )
        {
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        else if( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() )
        {
            tempField = acStrToStr( fildOption[0][index] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
    }
    
    //Fermer le fichier excel
    book->release();
    return lastField;
}

vector<AcString> createCurveField( const AcString& paramFile )
{

    //Listes des champs pour les blocks
    vector<AcString> field;
    vector<AcString> lastField;
    
    field.push_back( _T( "HANDLE" ) );
    field.push_back( _T( "LAYER" ) );
    field.push_back( _T( "COLOR" ) );
    field.push_back( _T( "LINETYPE" ) );
    field.push_back( _T( "LTSCALE" ) );
    field.push_back( _T( "LWEIGHT" ) );
    field.push_back( _T( "TRANSPARENCY" ) );
    field.push_back( _T( "THICKNESS" ) );
    field.push_back( _T( "MATERIAL" ) );
    field.push_back( _T( "LENGTH2D" ) );
    field.push_back( _T( "LENGTH3D" ) );
    field.push_back( _T( "STATUS" ) );
    
    
    //Charcger le fichier de configuration
    Book* book = xlCreateBook();
    
    vector<AcString> fildOption[2];
    
    if( book->load( paramFile ) )
    {
        Sheet* sheet = book->getSheet( 0 );
        
        if( sheet )
            for( int row = sheet->firstRow(); row < sheet->lastRow(); ++row )
                for( int col = sheet->firstCol(); col < sheet->lastCol(); col++ )
                    fildOption[col].push_back( sheet->readStr( row, col ) );
    }
    
    else
    {
    
        for( int i = 0; i < field.size(); i++ )
        {
            string tempField = acStrToStr( field[i] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        return lastField;
    }
    
    vector<AcString>::iterator it = fildOption[1].begin();
    
    for( it; it != fildOption[1].end(); ++it )
    {
        //Recuperer le numero de la ligne
        int index = it - fildOption[1].begin();
        string tempField = acStrToStr( *it );
        
        if( ( it->compareNoCase( "" ) != 0 ) && ( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() ) )
        {
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        else if( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() )
        {
            tempField = acStrToStr( fildOption[0][index] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
    }
    
    //Fermer le fichier excel
    book->release();
    return lastField;
}


vector<AcString> createHatchField( const AcString& paramFile )
{
    //Listes des champs pour les blocks
    vector<AcString> field;
    vector<AcString> lastField;
    
    field.push_back( _T( "HANDLE" ) );
    field.push_back( _T( "LAYER" ) );
    field.push_back( _T( "COLOR" ) );
    field.push_back( _T( "LINETYPE" ) );
    field.push_back( _T( "LTSCALE" ) );
    field.push_back( _T( "LWEIGHT" ) );
    field.push_back( _T( "TRANSPARENCY" ) );
    field.push_back( _T( "MATERIAL" ) );
    field.push_back( _T( "POSX" ) );
    field.push_back( _T( "POSY" ) );
    field.push_back( _T( "POSZ" ) );
    field.push_back( _T( "AREA" ) );
    field.push_back( _T( "PATTERN" ) );
    
    
    //Charcger le fichier de configuration
    Book* book = xlCreateBook();
    
    vector<AcString> fildOption[2];
    
    if( book->load( paramFile ) )
    {
        Sheet* sheet = book->getSheet( 0 );
        
        if( sheet )
            for( int row = sheet->firstRow(); row < sheet->lastRow(); ++row )
                for( int col = sheet->firstCol(); col < sheet->lastCol(); col++ )
                    fildOption[col].push_back( sheet->readStr( row, col ) );
    }
    
    else
    {
    
        for( int i = 0; i < field.size(); i++ )
        {
            string tempField = acStrToStr( field[i] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        return lastField;
    }
    
    vector<AcString>::iterator it = fildOption[1].begin();
    
    for( it; it != fildOption[1].end(); ++it )
    {
        //Recuperer le numero de la ligne
        int index = it - fildOption[1].begin();
        string tempField = acStrToStr( *it );
        
        if( ( it->compareNoCase( "" ) != 0 ) && ( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() ) )
        {
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        else if( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() )
        {
            tempField = acStrToStr( fildOption[0][index] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
    }
    
    //Fermer le fichier excel
    book->release();
    return lastField;
}

void drawPointShp( AcDbPoint* point,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field )
{
    double* pointX;
    double* pointY;
    double* pointZ;
    
    SHPObject* object;
    
    //Recuperer le point d'insertion du block
    AcGePoint3d pointPosition = point->position();
    
    pointX = &pointPosition.x;
    pointY = &pointPosition.y;
    pointZ = &pointPosition.z;
    
    object = SHPCreateObject( SHPT_POINTZ, -1, 1, NULL, NULL, 1, pointX, pointY, pointZ, NULL );
    int iShape = SHPWriteObject( myShapeFile, -1, object );
    
    //Ecrire le fichier dfg
    writePoint( point, dbfHandle, iShape, field );
    
    //Fermer la shapefile
    SHPDestroyObject( object );
}


void drawPonctuel( AcDbEntity* ent,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field )
{
    double* pointX;
    double* pointY;
    double* pointZ;
    
    SHPObject* object;
    
    //Recuperer le point d'insertion du block
    AcGePoint3d pos_entity;
    
    if( ent->isKindOf( AcDbPoint::desc() ) )
    {
        //Caster l'entité
        AcDbPoint* ln = AcDbPoint::cast( ent );
        
        //Recuperer la position de l'entité
        pos_entity = ln->position();
    }
    
    else if( ent->isKindOf( AcDbText::desc() ) )
    {
        //Caster l'entité
        AcDbText* ln = AcDbText::cast( ent );
        
        //Recuperer la position de l'entité
        pos_entity = ln->position();
    }
    
    else if( ent->isKindOf( AcDbCircle::desc() ) )
    {
        //Caster l'entité
        AcDbCircle* ln = AcDbCircle::cast( ent );
        
        //Recuperer la position de l'entité
        pos_entity = ln->center();
    }
    
    else if( ent->isKindOf( AcDbBlockReference::desc() ) )
    {
        //Caster l'entité
        AcDbBlockReference* ln = AcDbBlockReference::cast( ent );
        
        //Recuperer la position de l'entité
        pos_entity = ln->position();
    }
    
    else if( ent->isKindOf( AcDbMText::desc() ) )
    {
        //Caster l'entité
        AcDbMText* ln = AcDbMText::cast( ent );
        
        //Recuperer la position de l'entité
        pos_entity = ln->location();
    }
    
    pointX = &pos_entity.x;
    pointY = &pos_entity.y;
    pointZ = &pos_entity.z;
    
    object = SHPCreateObject( SHPT_POINTZ, -1, 1, NULL, NULL, 1, pointX, pointY, pointZ, NULL );
    int iShape = SHPWriteObject( myShapeFile, -1, object );
    
    //Ecrire le fichier dfg
    writePoint( ent, dbfHandle, iShape, field );
    
    //Fermer la shapefile
    SHPDestroyObject( object );
}


void drawCurve( AcDbEntity* ent,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field )
{
    double* pointX;
    double* pointY;
    double* pointZ;
    
    SHPObject* object;
    
    //Recuperer le point d'insertion du block
    AcGePoint3d pos_entity;
    
    if( ent->isKindOf( AcDbLine::desc() ) )
    {
        //Caster l'entité
        AcDbLine* line = AcDbLine::cast( ent );
        
        pointX = new double[2];
        pointY = new double[2];
        pointZ = new double[2];
        
        pointX[0] = line->startPoint().x;
        pointY[0] = line->startPoint().y;
        pointZ[0] = line->startPoint().z;
        
        pointX[1] = line->endPoint().x;
        pointY[1] = line->endPoint().y;
        pointZ[1] = line->endPoint().z;
        
        
        if( ( pointX != NULL ) && ( pointY != NULL ) && ( pointZ != NULL ) )
        {
            SHPObject* object = SHPCreateObject( SHPT_ARCZ, -1, 1, NULL, NULL, 2, pointX, pointY, pointZ, NULL );
            int iShape = SHPWriteObject( myShapeFile, -1, object );
            
            //Fermer la shapefile
            SHPDestroyObject( object );
            
            //Ecrire le fichier dfg
            writeCurve( line, dbfHandle, iShape, field );
        }
        
    }
    
    else if( ent->isKindOf( AcDbPolyline::desc() ) )
    {
        //Caster l'entité
        AcDbPolyline* poly2d = AcDbPolyline::cast( ent );
        
        int vertNumber = poly2d->numVerts();
        
        if( poly2d->isClosed() )
            vertNumber++;
            
        pointX = new double[vertNumber];
        pointY = new double[vertNumber];
        pointZ = new double[vertNumber];
        
        //Boucler sur les sommets
        for( int j = 0; j < poly2d->numVerts(); j++ )
        {
            AcGePoint3d pt;
            poly2d->getPointAt( j, pt );
            pointX[j] = pt.x;
            pointY[j] = pt.y;
            pointZ[j] = pt.z;
        }
        
        if( poly2d->isClosed() )
        {
            AcGePoint3d pt;
            poly2d->getPointAt( 0, pt );
            pointX[vertNumber - 1] = pt.x;
            pointY[vertNumber - 1] = pt.y;
            pointZ[vertNumber - 1] = pt.z;
        }
        
        if( ( pointX != NULL ) && ( pointY != NULL ) && ( pointZ != NULL ) )
        {
            SHPObject* object = SHPCreateObject( SHPT_ARCZ, -1, 1, NULL, NULL, vertNumber, pointX, pointY, pointZ, NULL );
            int iShape = SHPWriteObject( myShapeFile, -1, object );
            //Fermer la shapefile
            SHPDestroyObject( object );
            //Ecrire le fichier dfg
            writeCurve( poly2d, dbfHandle, iShape, field );
        }
        
    }
    
    else if( ent->isKindOf( AcDbCircle::desc() ) )
    {
        //Caster l'entité
        AcDbCircle* circle = AcDbCircle::cast( ent );
        
        //Recuperer tous les points du cercle
        AcGePoint3dArray arrPoint;
        
        double fleche;
        
        arrPoint = discretize( circle, 0.00001 );
        
        //Recuperer le premier element du tableau
        AcGePoint3d pt = arrPoint[0];
        
        //Ajouter en fin du tableau le contenu du premier indice du tableau
        arrPoint.append( pt );
        
        //Recuperer la taille du tableau
        int vertNumber = arrPoint.size();
        
        pointX = new double[vertNumber];
        pointY = new double[vertNumber];
        pointZ = new double[vertNumber];
        
        //Boucler sur les sommets
        for( int j = 0; j < vertNumber; j++ )
        {
            AcGePoint3d pt = arrPoint[j];
            pointX[j] = pt.x;
            pointY[j] = pt.y;
            pointZ[j] = pt.z;
        }
        
        if( ( pointX != NULL ) && ( pointY != NULL ) && ( pointZ != NULL ) )
        {
            SHPObject* object = SHPCreateObject( SHPT_ARCZ, -1, 1, NULL, NULL, vertNumber, pointX, pointY, pointZ, NULL );
            int iShape = SHPWriteObject( myShapeFile, -1, object );
            //Fermer la shapefile
            SHPDestroyObject( object );
            //Ecrire le fichier dfg
            writeCurve( circle, dbfHandle, iShape, field );
        }
    }
    
    else if( ent->isKindOf( AcDb3dPolyline::desc() ) )
    {
        //Caster l'entité
        AcDb3dPolyline* poly3d = AcDb3dPolyline::cast( ent );
        
        
        int vertNumber = getNumberOfVertex( poly3d );
        
        AcDbObjectIterator* iterPoly = poly3d->vertexIterator();
        AcDb3dPolylineVertex* vertex;
        int j = 0;
        
        if( poly3d->isClosed() )
            vertNumber++;
            
        pointX = new double[vertNumber];
        pointY = new double[vertNumber];
        pointZ = new double[vertNumber];
        
        
        for( iterPoly->start(); !iterPoly->done(); iterPoly->step() )
        {
            if( Acad::eOk == poly3d->openVertex( vertex, iterPoly->objectId(), AcDb::kForRead ) )
            {
                pointX[j] = vertex->position().x;
                pointY[j] = vertex->position().y;
                pointZ[j] = vertex->position().z;
            }
            
            j++;
        }
        
        vertex->close();
        
        if( poly3d->isClosed() )
        {
            AcGePoint3d p;
            poly3d->getStartPoint( p );
            pointX[vertNumber - 1] = p.x;
            pointY[vertNumber - 1] = p.y;
            pointZ[vertNumber - 1] = p.z;
        }
        
        
        if( ( pointX != NULL ) && ( pointY != NULL ) && ( pointZ != NULL ) )
        {
            SHPObject* object = SHPCreateObject( SHPT_ARCZ, -1, 1, NULL, NULL, vertNumber, pointX, pointY, pointZ, NULL );
            int iShape = SHPWriteObject( myShapeFile, -1, object );
            //Fermer la shapefile
            SHPDestroyObject( object );
            //Ecrire le fichier dfg
            writeCurve( poly3d, dbfHandle, iShape, field );
        }
        
    }
    
}

void drawHatchShp( AcDbHatch* hatch,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field )
{
    double* pointX;
    double* pointY;
    double* pointZ;
    
    SHPObject* object;
    
    AcDbExtents ext;
    
    hatch->getGeomExtents( ext );
    
    //Determiner la position du centre
    AcGePoint3d pointPosition = midPoint3d( ext.minPoint(), ext.maxPoint() );
    
    pointX = &pointPosition.x;
    pointY = &pointPosition.y;
    pointZ = &pointPosition.z;
    
    object = SHPCreateObject( SHPT_POINTZ, -1, 1, NULL, NULL, 1, pointX, pointY, pointZ, NULL );
    int iShape = SHPWriteObject( myShapeFile, -1, object );
    
    //Ecrire le fichier dfg
    writeHatch( hatch, dbfHandle, iShape, field );
    
    //Fermer la shapefile
    SHPDestroyObject( object );
}

/* ==> END POINT SHAPE <== */


/* ==>  BLOCK SHAPE <== */

void drawBlockShp( AcDbBlockReference* block,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field )
{
    double* pointX;
    double* pointY;
    double* pointZ;
    
    SHPObject* object;
    
    //Recuperer les attrubtus
    map<AcString, AcString> att = getBlockAttWithValuesList( block );
    
    //Rajouter le chap du dbf
    map<AcString, AcString>::iterator it = att.begin();
    
    while( it != att.end() )
    {
        if( find( field.begin(), field.end(), it->first ) == field.end() )
        {
            DBFAddField( dbfHandle, it->first, FTString, MAXCHAR, 0 );
            field.push_back( it->first );
        }
        
        it++;
    }
    
    //Recuperer le point d'insertion du block
    AcGePoint3d blockInsertPoint = block->position();
    
    pointX = &blockInsertPoint.x;
    pointY = &blockInsertPoint.y;
    pointZ = &blockInsertPoint.z;
    
    object = SHPCreateObject( SHPT_POINTZ, -1, 1, NULL, NULL, 1, pointX, pointY, pointZ, NULL );
    int iShape = SHPWriteObject( myShapeFile, -1, object );
    
    //Ecrire le fichier dfg
    writeBlock( block, dbfHandle, iShape, field );
    
    //Fermer la shapefile
    SHPDestroyObject( object );
    
    //Fermer le block
    block->close();
}

vector<AcString> createBlockField( const AcString& paramFile )
{
    //Listes des champs pour les blocks
    vector<AcString> field;
    vector<AcString> lastField;
    field.push_back( _T( "HANDLE" ) );
    field.push_back( _T( "LAYER" ) );
    field.push_back( _T( "COLOR" ) );
    field.push_back( _T( "LINETYPE" ) );
    field.push_back( _T( "LTSCALE" ) );
    field.push_back( _T( "LWEIGHT" ) );
    field.push_back( _T( "TRANSPARENCY" ) );
    field.push_back( _T( "BLOCK" ) );
    //field.push_back( _T( "ANGLE" ) );
    field.push_back( _T( "MATERIAL" ) );
    field.push_back( _T( "POSX" ) );
    field.push_back( _T( "POSY" ) );
    field.push_back( _T( "POSZ" ) );
    field.push_back( _T( "ROTATION" ) );
    
    //Charcger le fichier de configuration
    Book* book = xlCreateBook();
    
    vector<AcString> fildOption[2];
    
    if( book->load( paramFile ) )
    {
        Sheet* sheet = book->getSheet( 0 );
        
        if( sheet )
            for( int row = sheet->firstRow(); row < sheet->lastRow(); ++row )
                for( int col = sheet->firstCol(); col < sheet->lastCol(); col++ )
                    fildOption[col].push_back( sheet->readStr( row, col ) );
    }
    
    else
    {
    
        for( int i = 0; i < field.size(); i++ )
        {
            string tempField = acStrToStr( field[i] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        return lastField;
    }
    
    vector<AcString>::iterator it = fildOption[1].begin();
    
    for( it; it != fildOption[1].end(); ++it )
    {
        //Recuperer le numero de la ligne
        int index = it - fildOption[1].begin();
        string tempField = acStrToStr( *it );
        
        if( ( it->compareNoCase( "" ) != 0 ) && ( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() ) )
        {
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        else if( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() )
        {
            tempField = acStrToStr( fildOption[0][index] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
    }
    
    //Fermer le fichier excel
    book->release();
    return lastField;
}


vector<AcString> createBlockPoly2dField( const AcString& paramFile )
{
    //Listes des champs pour les blocks
    vector<AcString> field;
    vector<AcString> lastField;
    field.push_back( _T( "HANDLE" ) );
    field.push_back( _T( "LAYER" ) );
    field.push_back( _T( "COLOR" ) );
    field.push_back( _T( "LINETYPE" ) );
    field.push_back( _T( "LTSCALE" ) );
    field.push_back( _T( "LWEIGHT" ) );
    field.push_back( _T( "TRANSPARENCY" ) );
    field.push_back( _T( "BLOCK" ) );
    //field.push_back( _T( "ANGLE" ) );
    field.push_back( _T( "MATERIAL" ) );
    field.push_back( _T( "X" ) );
    field.push_back( _T( "Y" ) );
    field.push_back( _T( "Z" ) );
    field.push_back( _T( "ROTATION" ) );
    field.push_back( _T( "POLYHANDLE" ) );
    field.push_back( _T( "PLENGTH2D" ) );
    field.push_back( _T( "PLENGTH3D" ) );
    field.push_back( _T( "POLYAREA" ) );
    
    //Charcger le fichier de configuration
    Book* book = xlCreateBook();
    
    vector<AcString> fildOption[2];
    
    if( book->load( paramFile ) )
    {
        Sheet* sheet = book->getSheet( 0 );
        
        if( sheet )
            for( int row = sheet->firstRow(); row < sheet->lastRow(); ++row )
                for( int col = sheet->firstCol(); col < sheet->lastCol(); col++ )
                    fildOption[col].push_back( sheet->readStr( row, col ) );
    }
    
    else
    {
    
        for( int i = 0; i < field.size(); i++ )
        {
            string tempField = acStrToStr( field[i] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        return lastField;
    }
    
    vector<AcString>::iterator it = fildOption[1].begin();
    
    for( it; it != fildOption[1].end(); ++it )
    {
        //Recuperer le numero de la ligne
        int index = it - fildOption[1].begin();
        string tempField = acStrToStr( *it );
        
        if( ( it->compareNoCase( "" ) != 0 ) && ( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() ) )
        {
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        else if( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() )
        {
            tempField = acStrToStr( fildOption[0][index] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
    }
    
    //Fermer le fichier excel
    book->release();
    return lastField;
}


vector<AcString> createBlockPoly3dField( const AcString& paramFile )
{
    //Listes des champs pour les blocks
    vector<AcString> field;
    vector<AcString> lastField;
    field.push_back( _T( "HANDLE" ) );
    field.push_back( _T( "LAYER" ) );
    field.push_back( _T( "COLOR" ) );
    field.push_back( _T( "LINETYPE" ) );
    field.push_back( _T( "LTSCALE" ) );
    field.push_back( _T( "LWEIGHT" ) );
    field.push_back( _T( "TRANSPARENCY" ) );
    field.push_back( _T( "BLOCK" ) );
    //field.push_back( _T( "ANGLE" ) );
    field.push_back( _T( "MATERIAL" ) );
    field.push_back( _T( "POSX" ) );
    field.push_back( _T( "POSY" ) );
    field.push_back( _T( "POSZ" ) );
    field.push_back( _T( "ROTATION" ) );
    field.push_back( _T( "POLYHANDLE" ) );
    field.push_back( _T( "PLENGTH2D" ) );
    field.push_back( _T( "LENGTH3D" ) );
    field.push_back( _T( "POLYAREA" ) );
    
    //Charcger le fichier de configuration
    Book* book = xlCreateBook();
    
    vector<AcString> fildOption[2];
    
    if( book->load( paramFile ) )
    {
        Sheet* sheet = book->getSheet( 0 );
        
        if( sheet )
            for( int row = sheet->firstRow(); row < sheet->lastRow(); ++row )
                for( int col = sheet->firstCol(); col < sheet->lastCol(); col++ )
                    fildOption[col].push_back( sheet->readStr( row, col ) );
    }
    
    else
    {
    
        for( int i = 0; i < field.size(); i++ )
        {
            string tempField = acStrToStr( field[i] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        return lastField;
    }
    
    vector<AcString>::iterator it = fildOption[1].begin();
    
    for( it; it != fildOption[1].end(); ++it )
    {
        //Recuperer le numero de la ligne
        int index = it - fildOption[1].begin();
        string tempField = acStrToStr( *it );
        
        if( ( it->compareNoCase( "" ) != 0 ) && ( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() ) )
        {
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
        
        else if( find( field.begin(), field.end(), fildOption[0][index] ) != field.end() )
        {
            tempField = acStrToStr( fildOption[0][index] );
            
            for( int i = 10; i < tempField.length(); )
                tempField.pop_back();
                
            lastField.push_back( strToAcStr( tempField ) );
        }
    }
    
    //Fermer le fichier excel
    book->release();
    return lastField;
}

void drawblocSurfacekShp( AcDbBlockReference* block,
    vector<AcGePoint3d> pt,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field )
{
    double* pointX;
    double* pointY;
    double* pointZ;
    
    int vertNumber = pt.size();
    
    pointX = new double[vertNumber];
    pointY = new double[vertNumber];
    pointZ = new double[vertNumber];
    
    SHPObject* object;
    
    //Recuperer les attrubtus
    map<AcString, AcString> att = getBlockAttWithValuesList( block );
    
    //Rajouter le chap du dbf
    map<AcString, AcString>::iterator it = att.begin();
    
    while( it != att.end() )
    {
        if( find( field.begin(), field.end(), it->first ) == field.end() )
        {
            DBFAddField( dbfHandle, it->first, FTString, MAXCHAR, 0 );
            field.push_back( it->first );
        }
        
        it++;
    }
    
    int i = 0;
    
    //Boucler sur la listes des points
    for( AcGePoint3d p : pt )
    {
        pointX[i] = p.x;
        pointY[i] = p.y;
        pointZ[i] = p.z;
        i++;
    }
    
    if( ( pointX != NULL ) && ( pointY != NULL ) && ( pointZ != NULL ) )
    {
        SHPObject* object = SHPCreateObject( SHPT_POLYGONZ, -1, 1, NULL, NULL, vertNumber, pointX, pointY, pointZ, NULL );
        int iShape = SHPWriteObject( myShapeFile, -1, object );
        //Fermer la shapefile
        SHPDestroyObject( object );
        
        //Ecrire le fichier dfg
        writeBlock( block, dbfHandle, iShape, field );
    }
}


void drawBlockPoly( AcDbBlockReference* block,
    AcDbEntity* polyEnt,
    vector<AcGePoint3d> pt,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field )
{
    double* pointX;
    double* pointY;
    double* pointZ;
    
    int vertNumber = pt.size();
    
    pointX = new double[vertNumber];
    pointY = new double[vertNumber];
    pointZ = new double[vertNumber];
    
    SHPObject* obblock;
    
    
    //Recuperer les attrubtus
    map<AcString, AcString> att = getBlockAttWithValuesList( block );
    
    //Rajouter le chap du dbf
    map<AcString, AcString>::iterator it = att.begin();
    
    while( it != att.end() )
    {
        if( find( field.begin(), field.end(), it->first ) == field.end() )
        {
            DBFAddField( dbfHandle, it->first, FTString, MAXCHAR, 0 );
            field.push_back( it->first );
        }
        
        it++;
    }
    
    int i = 0;
    
    //Boucler sur la listes des points
    for( AcGePoint3d p : pt )
    {
        pointX[i] = p.x;
        pointY[i] = p.y;
        pointZ[i] = p.z;
        i++;
    }
    
    if( ( pointX != NULL ) && ( pointY != NULL ) && ( pointZ != NULL ) )
    {
        SHPObject* object = SHPCreateObject( SHPT_POLYGONZ, -1, 1, NULL, NULL, vertNumber, pointX, pointY, pointZ, NULL );
        
        //SHPObject* object = SHPCreateObject( SHPT_ARCZ, -1, 1, NULL, NULL, vertNumber, pointX, pointY, pointZ, NULL );
        
        //
        int iShape = SHPWriteObject( myShapeFile, -1, object );
        
        //Liberer la mémoire
        SHPDestroyObject( object );
        
        //Ecrire le fichier dfg
        writeBlockPoly( block, polyEnt, dbfHandle, iShape, field );
    }
    
}



void drawblocSurfacekShp( AcDbText* txt,
    vector<AcGePoint3d> pt,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field )
{
    double* pointX;
    double* pointY;
    double* pointZ;
    
    int vertNumber = pt.size();
    
    pointX = new double[vertNumber];
    pointY = new double[vertNumber];
    pointZ = new double[vertNumber];
    
    SHPObject* object;
    
    int i = 0;
    
    //Boucler sur la listes des points
    for( AcGePoint3d p : pt )
    {
        pointX[i] = p.x;
        pointY[i] = p.y;
        pointZ[i] = p.z;
        i++;
    }
    
    if( ( pointX != NULL ) && ( pointY != NULL ) && ( pointZ != NULL ) )
    {
        SHPObject* object = SHPCreateObject( SHPT_POLYGONZ, -1, 1, NULL, NULL, vertNumber, pointX, pointY, pointZ, NULL );
        int iShape = SHPWriteObject( myShapeFile, -1, object );
        //Fermer la shapefile
        SHPDestroyObject( object );
        //Ecrire le fichier dfg
        writeText( txt, dbfHandle, iShape, field );
    }
}

void drawTextPoly(
    AcDbText* txt,
    AcDbEntity* id,
    vector<AcGePoint3d> pt,
    SHPHandle& myShapeFile,
    DBFHandle& dbfHandle,
    vector<AcString>& field )
{
    double* pointX;
    double* pointY;
    double* pointZ;
    
    int vertNumber = pt.size();
    
    pointX = new double[vertNumber];
    pointY = new double[vertNumber];
    pointZ = new double[vertNumber];
    
    SHPObject* object;
    
    int i = 0;
    
    //Boucler sur la listes des points
    for( AcGePoint3d p : pt )
    {
        pointX[i] = p.x;
        pointY[i] = p.y;
        pointZ[i] = p.z;
        i++;
    }
    
    if( ( pointX != NULL ) && ( pointY != NULL ) && ( pointZ != NULL ) )
    {
        SHPObject* object = SHPCreateObject( SHPT_POLYGONZ, -1, 1, NULL, NULL, vertNumber, pointX, pointY, pointZ, NULL );
        int iShape = SHPWriteObject( myShapeFile, -1, object );
        
        //Fermer la shapefile
        SHPDestroyObject( object );
        
        //Ecrire le fichier dfg
        writeTextPoly( txt, id, dbfHandle, iShape, field );
    }
}

/* ==> END BLOCK SHAPE  <== */


void writePrjOrNot( const AcString& filePath )
{
    //Demander si l'utilisateur veut ajouter un prj ou pas
    bool bWrite = true;
    
    if( RTCAN == askYesOrNo( bWrite, _T( "Ajouter le prj?" ) ) )
    {
        print( "Export sans prj" );
        return;
    }
    
    if( !bWrite )
    {
        print( "Export sans prj" );
        return;
    }
    
    //Entrer le ESPG
    int epsg = 2000;
    AcString pr;
    
    // On demande au dessinateur de rentrer une valeur
    pr.format( L"%s <%d>: ", _T( "EPSG :" ), epsg );
    
    if( RTCAN == acedGetInt( pr, &epsg ) )
    {
        print( "Export sans prj" );
        return;
    }
    
    //Lire le fichier excel
    string EPSGFILE = "C:\\Futurmap\\Outils\\GStarCAD2020\\CFG\\EPSG.xlsx" ;
    
    if( !isFileExisting( EPSGFILE ) )
    {
        print( "Listes des EPSG introuvables" );
        return;
    }
    
    //Lire le fichier excel
    xlnt::workbook epsgfile;
    epsgfile.load( latin1_to_utf8( EPSGFILE ) );
    
    xlnt::worksheet sheet1 = epsgfile.active_sheet();
    
    //Bouleen pour voir le EPSG est trouve ou pas
    int bFound = INT_MIN;
    
    //Boucler sur les lignes
    int iLine =  sheet1.highest_row() - sheet1.lowest_row();
    
    for( int i = 1; i < iLine; i++ )
    {
        try
        {
            //Recuperer le numero de la EPSG
            string sCell = sheet1.cell( 1, i ).to_string();
            
            if( !isAcstringNumber( strToAcStr( sCell ) ) )
                continue;
                
            //Recuperer le numero de EPSG
            int iEpsg = stoi( sCell );
            
            //si le numero de epsg est egal a celle entrer par l'utilisateur
            if( iEpsg == epsg )
            {
                bFound = i;
                break;
            }
        }
        
        catch( const std::exception& e )
        {
            print( e.what() );
            print( "Fichier EPSG corrompue" );
        }
    }
    
    //Si on a trouver l'index de la ligne du numero de EPSG
    if( bFound != INT_MIN )
    {
        string sref = sheet1.cell( 2, bFound ).to_string();
        
        //Chemin du prj
        AcString prj2Path = filePath;
        
        //Ecrire le prj
        if( prj2Path.find( _T( ".shp" ) ) != -1 )
            prj2Path.replace( ".shp", ".prj" );
        else prj2Path.append( _T( ".prj" ) );
        
        string sPrj = acStrToStr( prj2Path );
        
        std::ofstream prjFile( sPrj );
        
        prjFile << sref;
        prjFile.close();
    }
    
    else
        print( "Numero introuvable" );
}