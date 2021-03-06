import { Component, OnInit } from '@angular/core';
import * as fs from "fs";
import { FormControl } from '@angular/forms';
import { saveAs } from "file-saver";
import { Document, Table, TableRow, WidthType, BorderStyle, TableCell, AlignmentType, Packer, Paragraph, TextRun, HeadingLevel, ShadingType, SectionType, ImageRun } from "docx";

@Component({
  selector: 'app-generator-report',
  templateUrl: './generator-report.component.html',
  styleUrls: ['./generator-report.component.css']
})
export class GeneratorReportComponent {

  constructor() { }

  //     let rawdata = fs.readFileSync('Data.json').toString("utf8");
  // let expert = JSON.parse(rawdata);
  
  // Documents contain sections, you can have multiple sections per document, go here to learn more about sections
  // This simple example will only contain one section
  // public getVal(val) : string{
  //     var a;
  //     a= console.warn(val);
  //     return a;
  // };


  public onSubmit(id) : string{
  var b= "";
  b = (<HTMLInputElement>document.getElementById(id)).value;
  return b}

  public onSubmit1(id) : string{
    var h= "";
    h = (<HTMLTextAreaElement>document.getElementById(id)).value;
    return h}
    public nature(id1, id2, id3, id4) :string { 
        var a =document.getElementById(id1) as HTMLInputElement;
        var b =document.getElementById(id1) as HTMLInputElement;
        var c =document.getElementById(id1) as HTMLInputElement;
        var d =document.getElementById(id1) as HTMLInputElement;
        var e="";

        if(a != null && a.checked) {
            e = "Terrain";
            return e;
        }
        else if(b != null && b.checked) {
            e = "Bureau";
            return e;
        }
        else if(c != null && c.checked) {
            e = "Immeuble";
            return e;
        }
        else if(d != null && d.checked) {
            e = "Commerce";
            return e;
        }
        else{
            return "";
        }


    }


  public display1(id3) :string { 
    var t =document.getElementById(id3) as HTMLInputElement;
    if(t != null && t.checked) {
        var c = "X";
        return c
    }
    else{
        return "";
    }
}
public display2(id2) :string { 
    var t =document.getElementById(id2) as HTMLInputElement;
    if(t != null && t.checked) {
        var c = "X";
        return c
    }
    else{
        return "";
    }
}
  public download(): void {
    const table = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("      ")],
                    }),
                    new TableCell({
                        children: [new Paragraph("OUI")],
                    }),
                    new TableCell({
                        children: [new Paragraph("NON")],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Plein centre")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("OUI"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("NON"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Proche centre (300)")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("centre"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("centre2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("P??riph??rie")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("per"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("per2"))
                        ],
                    }),
                ],
            }),


            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Terrain sur Rue")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("rue"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("rue2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Mitoyen ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("Mitoyen"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("Mitoyen2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Vis ?? Vis ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("vis"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("vis2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Espaces verts ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("vert"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("vert2"))
                        ],
                    }),
                ],
            }),
            


        ],
    });



    const table7 = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("  ")],
                    }),
                    new TableCell({
                        children: [new Paragraph("OUI ")],
                    }),
                    new TableCell({
                        children: [new Paragraph("Non ")],
                    }),
                    new TableCell({
                        children: [new Paragraph("OBSERVATIONS ")],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Remblais  ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1('Remblais'))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2('Remblais2'))
                        ],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph("    ")
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Pr??sence de cuves   ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1('cuve'))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2('cuve2'))
                        ],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph("    ")
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Anciennes activit?? poulluantes   ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1('activite'))],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display2('activite2'))],
                    }),
                    new TableCell({
                        children: [new Paragraph("     ")],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("prescriptions arch??o  ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("archeo"))],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display2('archeo2'))],
                    }),
                    new TableCell({
                        children: [new Paragraph("OBSERVATIONS ")],
                    }),
                ],
            }),
    
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Pr??sence de transformateur  ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1('trans'))],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display2('trans2'))],
                    }),
                    new TableCell({
                        children: [new Paragraph("  ")],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Inondation  ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1('ino'))],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display2('ino2'))],
                    }),
                    new TableCell({
                        children: [new Paragraph("   ")],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Reprises en sous oeuvres  ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1('ouvre'))],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display2('ouvre2'))],
                    }),
                    new TableCell({
                        children: [new Paragraph("  ")],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Pr??sence d'une voie ferr??  ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1('voi'))],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display2('voi2'))],
                    }),
                    new TableCell({
                        children: [new Paragraph("    ")],
                    }),
                ],
            }),

            
            
            

        ],
    });



    const table5 = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("      ")],
                    }),
                    new TableCell({
                        children: [new Paragraph("OUI")],
                    }),
                    new TableCell({
                        children: [new Paragraph("NON")],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Sonores")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("son"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("NON"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Proche centre (300)")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("centre"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("centre2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("P??riph??rie")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("per"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("per2"))
                        ],
                    }),
                ],
            }),


            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Terrain sur Rue")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("rue"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("rue2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Mitoyen ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("Mitoyen"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("Mitoyen2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Vis ?? Vis ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("vis"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("vis2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Espaces verts ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("vert"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("vert2"))
                        ],
                    }),
                ],
            }),
            


        ],
    });

    const nuissanceTable = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("      ")],
                    }),
                    new TableCell({
                        children: [new Paragraph("OUI")],
                    }),
                    new TableCell({
                        children: [new Paragraph("NON")],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Sonores")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("sono"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("sono2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Lignes ?? Haute tesion")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("ligne"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("ligne2"))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("visuelles")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("visu"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("visu2"))
                        ],
                    }),
                ],
            }),


            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Olfactives")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.display1("fact"))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.display2("fact2"))
                        ],
                    }),
                ],
            }),
            


        ],
    });
    

    const clientTable = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Client")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit('name'))],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Objet de l'evaluation")],
                    }),
                    new TableCell({
                        children: [new Paragraph("Valeur v??nale et marchande de l???actif "+ this.onSubmit('Pptedite'))],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Interlocuteur Client")],
                    }),
                    new TableCell({
                        children: [new Paragraph("Monsieur "+ this.onSubmit('name'))],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("PI??CES RECUPEREES AUPRES DE LA CONSERVATION ")],
                        rowSpan: 3,
                    }),
                    new TableCell({
                        children: [new Paragraph("Documents")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph("Statut")
                        ],
                    }),
                    new TableCell({
                        children: [
                            new Paragraph("Date de Document")
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Certificat de propri??t??")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit('certif'))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(" ")
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Status de Plan Cadastral")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit('splan'))],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(" ")
                        ],
                    }),
                ],
            }),

        ],
    });
    

    const table8 = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Date de la visite")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit('date'))],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Personnes present??es")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit('visite'))],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Commentaires")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit1('tr'))],
                    }),
                ],
            }),


        ],
    });
    





    const juridiqueTable = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Num??ro de TF ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit('TF'))],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Intitul?? de la propri??t?? ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('Pptedite'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Consistance ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph("  ")
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Conservation Fonci??re")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('city'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Propri??taire")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('name'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Superficie ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('surface'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Droit r??el ou charge fonci??re")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('reel'))
                        ],
                    }),
                ],
            }),


        ],
    });

    const table10 = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("P??rim??tre de l'??tude ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit1('perm'))],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Offre pr??sente sur le p??rim??tre ??tudi??")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('etd'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Barom??tre du march?? ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('baro'))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Diff??renciation du bien par rapport ?? l???offre")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('offre'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Prix de vente ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('prix'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("??volution & projets futurs de la zone")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('future'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Droit r??el ou charge fonci??re")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('reel'))
                        ],
                    }),
                ],
            }),


        ],
    });

    const table1 = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("NOM DU DOSSIER ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit('dossier'))],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Objet de La mission")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph("??valuation en valeur v??nale")
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Date d'expertise")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('exp'))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Adresse des actifs")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('actif'))
                        ],
                    }),
                ],
            }),


        ],
    });

    const table9 = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Cat??gories ")],
                    }),
                    new TableCell({
                        children: [new Paragraph("Caract??ristiques")],
                    }),
                    new TableCell({
                        children: [new Paragraph("Notes")],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("G??n??rale")],
                        rowSpan: 6,
                    }),
                    new TableCell({children: [new Paragraph("Orientation (bien traversant)") ],}),
                    new TableCell({children: [new Paragraph(this.onSubmit('or')) ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Pr??sence de place de parking")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('park'))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Etat des parties communes")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('des'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Organisation spatiale")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('spa'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Etat des rev??tements")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('rev'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("EAccessibilit?? du bien")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('ben'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Technique")],
                        rowSpan: 6,
                    }),
                    new TableCell({
                        children: [new Paragraph("Etanch??it??")],
                       
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('etan'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Etat de la boiserie")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('boi'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Etat des rev??tements de sol")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('sol'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Electricit??")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('do'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Totale")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph("   ")
                        ],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph("   ")
                        ],
                    }),
                ],
            }),



        ],
    });


    const table4 = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("          ")],
                    }),
                    new TableCell({
                        children: [new Paragraph("DISTANCE  ")],
                    }),
                    new TableCell({
                        children: [new Paragraph("OBSERVATIONS  ")],
                    }),
                ],
            }),


        ],
    });
//    const image = new ImageRun({
//     data: fs.readFileSync("/src/assets/img/Capture.PNG"),
//     transformation: {
//         width: 200,
//         height: 200,
//     }
// });

    const table3 = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Commerces ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit1('commerce'))],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Services ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('service'))
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Ecole")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('ecole'))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Transport")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('trsp'))
                        ],
                    }),
                ],
            }),


        ],
    });



    const table2 = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Commune ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit('commune'))],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Surface globale des biens (m??)")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit("surface"))
                        ],
                    }),
                ],
            }),
            

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Num??ro de titre foncier")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('TF'))
                        ],
                    }),
                ],
            }),


        ],
    });



    const situationTable = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Situation ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit('sit'))],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Objet d'usage  ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit("obj"))
                        ],
                    }),
                ],
            }),
            

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Caract??re de la zone ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit('zone'))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Evolution & Projection de la zone ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('prj'))
                        ],
                    }),
                ],
            }),

        ],
    });
    
    

    const proxTable = new Table({
        width: {
            size: 9070,
            type: WidthType.DXA,
        },
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Commerces ")],
                    }),
                    new TableCell({
                        children: [new Paragraph(this.onSubmit1('cmr'))],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Administrations publics  ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1("adm"))
                        ],
                    }),
                ],
            }),
            

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Ecoles  ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('col'))
                        ],
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph("Agences de services ")],
                    }),
                    new TableCell({
                        children: [
                            
                            new Paragraph(this.onSubmit1('srv'))
                        ],
                    }),
                ],
            }),

        ],
    });
    


    
  const doc = new Document({
      
      sections: [{
          properties: {
              type: SectionType.NEXT_PAGE,
          },
          children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                heading: HeadingLevel.TITLE,
                children: [
                    new TextRun({
                  text: "RAPPORT D???EXPERTISE",
                  bold: true,
                  italics: true,
                  color: '000000',
                  allCaps: true,
                  size: 50,
                  break: 1,
                  

                    }),
                    new TextRun({
                        text : "Propri??t?? dite "+ this.onSubmit('Pptedite'),
                        color : '000000',
                        size : 32,
                        break: 1
                    }),
                    new TextRun({
                        text : "Objet de titre foncier N??  "+ this.onSubmit('TF'),
                        color : '000000',
                        size : 32,
                        break: 1
                    }),
                    new TextRun({
                        text : "Situ?? ?? :  "+ this.onSubmit('Pptedite'),
                        color : '000000',
                        size : 32,
                        break: 1
                    }),
                  
                    // new ImageRun({
                    //     data:  readFileSync(this.onSubmit('image')) ,
                    //     transformation: {
                    //         width: 200,
                    //         height: 200,
                    //     },
                    // }),
                    
                    

                ],
                  spacing: {
                      before: 250,
                      after: 9000,
                  },
                  
              }),
              
              new Paragraph({
                  
                          text:"Sommaire",
                          
                          heading: HeadingLevel.TITLE,
                         
                         
                          
                          shading: {
                              type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                              color: "4682B4",
                              fill: "4682B4",
                          },
                          spacing: {
                              after: 200
                          }
                      }),
              new Paragraph({
                          text: "\t1-1.	La mission",
                        
                          heading: HeadingLevel.HEADING_1,
                         
                          spacing: {
                              after: 200
                          }
                      }),
              new Paragraph({
                          text: "\t1-2.	Certificat d?????valuation",
                         
                          heading: HeadingLevel.HEADING_1,
                          
                          spacing: {
                              after: 200
                          }
                      }),
              new Paragraph({
                          text: "\t1-3	DEFINITIONS",
                        
                          heading: HeadingLevel.HEADING_1,
                          
                          spacing: {
                              after: 500
                              
                          }
                      }),
                      new Paragraph({
                          text: "\t2-1.	DEFINITION DU BIEN",
                          
                          heading: HeadingLevel.HEADING_1,
                          
                          spacing: {
                              after: 200
                          }
                      }),
                      new Paragraph({
                          text: "\t2-2.	VISITE DU BIEN",
                         
                          heading: HeadingLevel.HEADING_1,
                          
                          spacing: {
                              after: 200
                          }
                      }),
              new Paragraph({
                          text: "\t2-3.	DESCRIPTION DU BIEN",
                          
                          heading: HeadingLevel.HEADING_1,
                          
                          spacing: {
                              after: 500
                              
                          }
                      }),
  
                      new Paragraph({
                          text: "\t3-1.	ETAT APPARENT DU BIEN",
                       
                          heading: HeadingLevel.HEADING_1,
                         
                          spacing: {
                              after: 200
                          }
                      }),
                      new Paragraph({
                          text: "\t3-2.	SITUATION FONCIER ??? JURIDIQUE",
                         
                          heading: HeadingLevel.HEADING_1,
                          
                          spacing: {
                              after: 200
                          }
                      }),
                      new Paragraph({
                          text: "\t3-3.	SITUATION URBANISTIQUE",
                         
                          heading: HeadingLevel.HEADING_1,
                          
                          spacing: {
                              after: 200
                          }
                      }),
              new Paragraph({
                          text: "\t3-4.	ETUDE DE MARCHE",
                       
                          heading: HeadingLevel.HEADING_1,
                          
                          spacing: {
                              after: 9000
                              
                          }
                      }),
                      new Paragraph({
                          text: "1. La Mission :   ",
                          
                          heading: HeadingLevel.HEADING_2,
                          
                          border: {
                              bottom: {
                                  color: "auto",
                                  space: 1,
                                  value: "single",
                                  size: 6,
                              },
                          },
                          spacing: {
                              after: 200
                              
                          }
                      }),
                      new Paragraph({
                          text: "Rapport d?????valuation du TF :"+ this.onSubmit('TF'),
                          heading: HeadingLevel.HEADING_2,
                         
                      
                          spacing: {
                              after: 200
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                                text: "Agissant suite ?? la demande de Monsieur  " + this.onSubmit("name")+" "  + +" en qualit?? de "+ this.onSubmit('qualite') +", tendant ?? d??finir la valeur v??nale des constructions de la propri??t?? dite "+this.onSubmit('Pptedite') + " objet du titre foncier  n?? "+ this.onSubmit('TF')+" sise : "+ this.onSubmit('adresse') + ". Nous proc??dons ?? une ??tude d?????valuation des actifs corporels (Terrain, b??timents , constructions,  installations et autres) qui nous ont ??t?? pr??sent??s et vous pr??sentons nos conclusions dans le pr??sent rapport.",
                                bold: false,
                                color: "000000",
                            }),
                            new TextRun({
                                text: ". Nous proc??dons ?? une ??tude d?????valuation des actifs corporels (Terrain, b??timents , constructions,  installations et autres) qui nous ont ??t?? pr??sent??s et vous pr??sentons nos conclusions dans le pr??sent rapport.",
                                bold: false,
                                color: "000000",
                            }),],
                       
                          heading: HeadingLevel.HEADING_3,
                          
                          spacing: {
                              after: 200
                              
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                          text: "\t1-	Enqu??te pr??alable aupr??s du requ??rant, notamment pour recueillir des informations sur les titres de la propri??t??, plans existants, certificat de propri??t??, ainsi que des ??l??ments comptables concernant les plantations, les constructions et les ??quipements.",
                          bold: false,
                          color: "000000",
                      }),],
                          heading: HeadingLevel.HEADING_3,
                          
                          spacing: {
                              after: 200
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                          text: "\t2-	Visite des lieux en date du"+this.onSubmit("date") +" pour appr??ciation des ??quipements et constructions r??ellement existantes, date de mise en place, avec quelques m??tr??s et enqu??te aupr??s des responsables sur place.",
                          bold: false,
                          color: "000000",
                      }),],
                          heading: HeadingLevel.HEADING_3,
                          
                          spacing: {
                              after: 200
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                          text: "\t3-	Consultation des donn??es informatiques (Google Maps) avec insertion des donn??es topographiques (coordonn??es Lambert). ",
                          bold: false,
                          color: "000000",
                      }),],
                          heading: HeadingLevel.HEADING_3,
                          
                          spacing: {
                              after: 200
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                          text: "\t4-	D??finition de chaque ??l??ment de la propri??t?? aux fins d???en estimer les valeurs.",
                          bold: false,
                          color: "000000",
                      }),],
                          heading: HeadingLevel.HEADING_3,
                          
                          spacing: {
                              after: 200
                              
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                          text: "\t5-	Enfin, nous avons men?? une ??tude pour d??finir la valeur v??nale de la propri??t??.",
                          bold: false,
                          color: "000000",
                      }),],
                          heading: HeadingLevel.HEADING_3,
                          // shading: {
                          //   type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                          //   color: "000000",
                          //   fill: "FFFFFF"},
                          spacing: {
                              after: 200
                              
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                          text: "\t6-	Nous utilisons la m??thode de "+this.onSubmit('mtd1') +", la m??thode de "+ this.onSubmit('mtd2')+" et la m??thode de " + this.onSubmit('mtd3'),
                          bold: false,
                          color: "000000",
                      }),],
                          heading: HeadingLevel.HEADING_3,
                          
                          spacing: {
                              after: 9000
                              
                          }
                      }),
                      new Paragraph({
                        text: "2. Certificat d?????valuation :",
                        heading: HeadingLevel.HEADING_2,
                        border: {
                            bottom: {
                                color: "auto",
                                space: 1,
                                value: "single",
                                size: 6,
                            },
                        },
                    
                        spacing: {
                            after: 200
                        }
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text:"A Monsieur "+ this.onSubmit("name"),
                          bold: false,
                          color: "000000",
                      }),],
                          heading: HeadingLevel.HEADING_3,
                          // shading: {
                          //   type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                          //   color: "000000",
                          //   fill: "FFFFFF"},
                          spacing: {
                              after: 200
                              
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                                text : "Objet : Estimation de la valeur d???un actif immobilier sis ?? "+ this.onSubmit("adresse"),
                          bold: true,
                          color: "000000",
                      }),],
                          heading: HeadingLevel.HEADING_3,
                          
                          spacing: {
                              after: 600
                              
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                                text: "Monsieur,",
                          bold: false,
                          color: "000000",
                      }),]}),
                      new Paragraph({
                        children: [
                      new TextRun({
                        text: "Nous avons l???honneur de vous communiquer notre estimation de la valeur d???un actif immobilier objet du TF 244104/07 Au regard des ??l??ments qui nous ont ??t?? communiqu??s, de notre ??valuation et des analyses d??crites dans notre pr??sent rapport d???expertise, nous retiendrons une Valeur V??nale (VV) de 737000 dhs (var : valeur venale)en date du 06-Feb-21pour (var : date rapport)cet actif libre de toutes occupations, charges ou   contrainte,",
                  bold: false,
                  color: "000000",
              }),
                    ],
                          heading: HeadingLevel.HEADING_3,
                          
                          spacing: {
                              after: 600
                              
                          }
                      }),
                      new Paragraph({
                        children: [
                            new TextRun({
                                text: "Nous restons ?? votre enti??re disposition pour vous apporter toute information qui vous semblerait n??cessaire sur le contenu du rapport d???expertise et la m??thodologie adopt??e.  \n",
                                bold: false,
                          color: "000000",
                      }),
                      new TextRun({
                          text: "Nous vous prions d???agr??er, Monsieur, l???expression de nos salutations distingu??es.",
                        bold: false,
                  color: "000000",
              }),
                    ],
                          heading: HeadingLevel.HEADING_3,
                          // shading: {
                          //   type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                          //   color: "000000",
                          //   fill: "FFFFFF"},
                          spacing: {
                              after: 300
                              
                          }
                      }),
                      new Paragraph({
                        text: "Salim BENMLIH  : Ing??nieur Expert Immobilier",
                        heading: HeadingLevel.HEADING_2,
                        spacing: {
                            after: 200
                        }
                    }),
                    new Paragraph({
                        text:"Mohammed Soussi Sadoq : Notaire Expert Immobilier RICS",
                        heading: HeadingLevel.HEADING_2,
                        spacing: {
                            after: 5000
                        }
                    }),
                    new Paragraph({
                        text :"3. D??finitions :       ",
                        heading: HeadingLevel.HEADING_2,
                        border: {
                            bottom: {
                                color: "auto",
                                space: 1,
                                value: "single",
                                size: 6,
                            },
                        },
                    
                        spacing: {
                            after: 200
                        }
                    }),
                    new Paragraph({
                        text : "Evaluation (Valuation) : Avis ??crit d???un expert sur la valeur des droits li??s ?? un actif immobilier, ?? la date d?????valuation. L?????valuation sera ??tablie, sauf limites convenues dans les conditions du contrat, ?? la suite d???une visite, et de toute investigation ou enqu??te compl??mentaire pertinente en conformit?? avec la nature de l???actif immobilier et les objectifs d??finis dans le contrat d?????valuation.  ",
                        heading: HeadingLevel.HEADING_3,
                        spacing: {
                            after: 200
                        }
                    }),
                    new Paragraph({
                        text : "Immeuble (Property) : Ensemble des droits et int??r??ts portant sur des terrains (avec et sans constructions), installations techniques et ??quipements et tout actif pouvant se d??pr??cier sauf d??finition plus restrictive dict??e par le contexte. Cette expression s???applique ??galement aux autres actifs d??tenus tels que des stocks ou travaux en cours, lorsque l?????valuation a pour but d???int??grer une valeur dans les ??tats financiers.  ",
                        heading: HeadingLevel.HEADING_3,
                        spacing: {
                            after: 200
                        }
                    }),
                    new Paragraph({
                        text: "Juste Valeur (Fair Value) : Montant pour lequel un actif peut ??tre ??chang?? ou une dette r??gl??e entre des parties bien inform??es, consentantes et agissant dans des conditions de concurrence normales.  ",
                        heading: HeadingLevel.HEADING_3,
                        spacing: {
                            after: 200
                        }
                    }),
                    new Paragraph({
                        text: "M??thode d?????valuation (Method of Valuation) : Approche ou technique employ??e pour calculer la valeur d??finie selon le cadre d?????valuation.  ",
                        heading: HeadingLevel.HEADING_3,
                        spacing: {
                            after: 200
                        }
                    }),
                    new Paragraph({
                        text: "Valeur v??nale (VV) Market Value :  Somme d???argent estim??e contre laquelle un immeuble serait ??chang?? ?? la date d?????valuation entre un acheteur consentant et un vendeur consentant dans une transaction ??quilibr??e apr??s une commercialisation ad??quate o?? les parties ont l???une et l???autre agit en toute connaissance, prudemment et sans pression.     ",
                        heading: HeadingLevel.HEADING_3,
                        spacing: {
                            after: 200
                        }
                    }),
                    new Paragraph({
                        text: "Source : NORMES D???EVALUATION DE LA RICS (sixi??me ??dition, version Fran??aise, avril 2010)   ",
                        heading: HeadingLevel.HEADING_3,
                        spacing: {
                            after: 1900
                        }
                    }),
                    new Paragraph({
                  
                        children: [
                            new TextRun({
                                text: "       I - Le site ",
                          bold: false,
                          color: "000000",
                      }),],
                        
                        heading: HeadingLevel.HEADING_1,
                       
                       
                        
                        shading: {
                            type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                            color: "4682B4",
                            fill: "4682B4",
                        },
                        spacing: {
                            after: 200
                        }
                    }),
                    table1,
                    table2,
                    new Paragraph({
                  
                        
                        children: [
                            new TextRun({
                                text:"     SITUATION               ACCESSIBILITE (Distance en m2)",
                                bold: false,
                          color: "000000",
                      }),],
                        
                        heading: HeadingLevel.HEADING_1,
                       
                        
                        shading: {
                            type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                            color: "4682B4",
                            fill: "4682B4",
                        },
                        spacing: {
                            after: 200
                        }
                    }),
                    table,
                    new Paragraph({
                  
                        children: [
                            new TextRun({
                                text: "       Observations ",
                          bold: false,
                          color: "000000",
                      }),],
                        
                        heading: HeadingLevel.HEADING_1,
                       
                       
                        
                        shading: {
                            type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                            color: "4682B4",
                            fill: "4682B4",
                        },
                        spacing: {
                            after: 200
                        }
                    }),
                    
                    table3,

                    new Paragraph({
                  
                        
                        children: [
                            new TextRun({
                                text:"     DESSERTE",
                                bold: true,
                          color: "000000",
                      }),],
                        
                        heading: HeadingLevel.HEADING_1,
                       
                        
                        shading: {
                            type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                            color: "4682B4",
                            fill: "4682B4",
                        },
                        spacing: {
                            after: 200
                        }
                    }),
                    table4,

                    new Paragraph({
                  
                        
                        children: [
                            new TextRun({
                                text:"     NUISSANCES",
                                bold: true,
                          color: "000000",
                      }),],
                        
                        heading: HeadingLevel.HEADING_1,
                       
                        
                        shading: {
                            type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                            color: "4682B4",
                            fill: "4682B4",
                        },
                        spacing: {
                            after: 200
                        }
                    }),
                    nuissanceTable,
                    new Paragraph({
                  
                        
                        children: [
                            new TextRun({
                                text:"     Risques",
                                bold: true,
                          color: "000000",
                      }),],
                        
                        heading: HeadingLevel.HEADING_1,
                       
                        
                        shading: {
                            type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                            color: "4682B4",
                            fill: "4682B4",
                        },
                        spacing: {
                            after: 200,
                            before : 20
                        }
                    }),
                    table7,
                    new Paragraph({children: [
                            new TextRun({
                                text:"     Observations Particuli??res",
                                bold: true,
                          color: "000000",
                      }),],
                        heading: HeadingLevel.HEADING_1,
                        shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                            color: "4682B4",
                            fill: "4682B4",},
                        spacing: {after: 200,
                            before : 20}}),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: this.onSubmit1('obs'),
                          bold: false,
                          color: "000000",
                      }),],
                      spacing: {
                        after: 200,
                        before : 20
                    }
                    }),
                    new Paragraph({children: [
                        new TextRun({
                            text:"     II - METHODOLOGIE D'EVALUATION ",
                            bold: true,
                      color: "000000",
                  }),],
                    heading: HeadingLevel.HEADING_1,
                    shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                        color: "4682B4",
                        fill: "4682B4",},
                    spacing: {after: 200,
                        before : 20}}),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Nos expertises sont r??alis??es conform??ment aux standards d?????valuation publi??s par Royal Institution of Chartered Surveyors (RICS).",
                              bold: false,
                              color: "000000",
                              break: 1
                          }),
                          new TextRun({
                              text: "Les m??thodologies d?????valuation pr??conis??es pour le bien expertis?? sont :",
                      bold: false,
                      color: "000000",
                      break: 1
                  }),

                  new TextRun({
                      text: this.onSubmit('mtd1'),
            bold: false,
            color: "000000",
            break: 1
        }),
        new TextRun({
            text: this.onSubmit('mtd2'),
  bold: false,
  color: "000000",
  break: 1
}),
new TextRun({
    text: this.onSubmit('mtd3'),
bold: false,
color: "000000",
break: 1
}),
                        ],
                          spacing: {
                            after: 200,
                            before : 20
                        }
                        }),

                        new Paragraph({children: [
                            new TextRun({
                                text:"     Hypoth??ses Particuli??res  ",
                                bold: true,
                          color: "000000",
                      }),],
                        heading: HeadingLevel.HEADING_1,
                        shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                            color: "4682B4",
                            fill: "4682B4",},
                        spacing: {after: 200,
                            before : 20}}),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: this.onSubmit1('hyp'),
                                  bold: false,
                                  color: "000000",
                              }),],
                              spacing: {
                                after: 200,
                                before : 20
                            }
                            }),
                            new Paragraph({children: [
                                new TextRun({
                                    text:"     III - COMPOSITION  ",
                                    bold: true,
                              color: "000000",
                          }),],
                            heading: HeadingLevel.HEADING_1,
                            shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                                color: "4682B4",
                                fill: "4682B4",},
                            spacing: {after: 200,
                                before : 20}}),
                                new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: "Nature de l'actif    : ",
                                      bold: true,
                                      color: "000000",
                                  }),
                                  new TextRun({
                                    text: this.nature("ter", "bur", "ime", "com"),
                              bold: false,
                              color: "000000",
                              break: 1
                          }),
                                ],
                                  spacing: {
                                    after: 200,
                                    before : 20
                                }
                                }), 
                                new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: "Surface de l'actif    :",
                                      bold: false,
                                      color: "000000",
                                      break: 1
                                  }),
                                ],
                                  spacing: {
                                    after: 300,
                                    before : 20
                                }
                                }), 
                                new Paragraph({children: [
                                    new TextRun({
                                        text:"     D??nomination     TF     Etage     Surface(m2)",
                                        bold: true,
                                  color: "000000",
                              }),],
                                heading: HeadingLevel.HEADING_1,
                                shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                                    color: "4682B4",
                                    fill: "4682B4",},
                                spacing: {after: 200,
                                    before : 20}}),
                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: "     " + this.onSubmit('Pptedite') +"     " +this.onSubmit('TF')+"     "+this.onSubmit('surface')+"     "+this.onSubmit('etage'),
                                          bold: false,
                                          color: "000000",
                                          break: 2
                                      }),
                                    ],
                                      spacing: {
                                        after: 300,
                                        before : 20
                                    }
                                    }), 

                                    new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: "Etat des lieux  :",
                                          bold: false,
                                          color: "000000",
                                          break: 1
                                      }),
                                      new TextRun({
                                        text: this.onSubmit('etat'),
                                  bold: false,
                                  color: "000000",
                                  break: 1
                              }),
                              new TextRun({
                                text: "Droit r??el ou Charge fonci??re : ",
                          bold: false,
                          color: "000000",
                          
                      }),
                      new TextRun({
                        text: this.onSubmit('droit'),
                  bold: false,
                  color: "000000",
                  break: 1
              }),
              new TextRun({
                text: "Etat Apparent :",
          bold: false,
          color: "000000",
          break: 1
      }),
      new TextRun({
        text: this.onSubmit('apparent'),
  bold: false,
  color: "000000",
  break: 1
}),
new TextRun({
    text: "Immeuble neuf : ",
bold: false,
color: "000000",
break: 1
}),
new TextRun({
    text: this.onSubmit('neuf'),
bold: false,
color: "000000",
break: 1
}),
new Paragraph({children: [
    new TextRun({
        text:"        IV - ETAT APPARENT dDU BIEN    ",
        bold: true,
  color: "000000",
}),],
heading: HeadingLevel.HEADING_1,
shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
    color: "4682B4",
    fill: "4682B4",},
spacing: {after: 200,
    before : 20}}),],
     }),
     new Paragraph({
        children: [
            new TextRun({
                text:"???	Une visite de l???actif a ??t?? effectu??e minutieusement afin de v??rifier son ??tat actuel et de s???informer de son environnement mitoyen. "+ this.onSubmit('apparent') +" . Un descriptif photographique a ??t?? fait en ce sens.",
          bold: false,
          color: "000000",
          break: 2
      }),
    ],
      spacing: {
        after: 300,
        before : 20
    }
    }),


    new Paragraph({
        alignment: AlignmentType.CENTER,
        heading: HeadingLevel.TITLE,
        children: [
            new TextRun({
          text: "CONTEXTE DE L'ETUDE ",
          bold: true,
          italics: true,
          color: '000000',
          allCaps: true,
          size: 50,
          break: 1,
          

            }),
        ],
        spacing: {
            before: 2500,
            after: 9000,
        },
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: "ans un but de pr??sentation r??elle et claire de la situation et de l?????valuation & la valorisation des actifs, nous avons opt?? pour une d??marche m??thodologique et distincte afin d?????tudier l???ensemble des aspects fondamen- taux de ces derniers.",
              bold: false,
              color: "000000",
              break: 1
          }),
          new TextRun({
              text:"Notre mission consistera en l?????valuation en valeur v??nale et marchande des biens selon les normes de la Royal Institution of Chartered Surveyors (RICS), ainsi qu???en l???estimation de la valeur v??nale. Notre d??marche s???articulera autour des actions suivantes afin d???affiner au mieux l?????valuation men??e :",
      bold: false,
      color: "000000",
      break: 1
  }),
  new TextRun({
text: "-    Une ??tude de l???environnement mitoyen et proche de l???actif ;",
    bold: false,
color: "000000",
break: 1
}),
new TextRun({
            text:"-    Une visite sur site ;",
        bold: false,
    color: "000000",
    break: 1
    }),
    new TextRun({
            text:"-    Une analyse des documents juridiques, urbanistiques et techniques de l???actif ;",
            bold: false,
        color: "000000",
        break: 1
        }),
        new TextRun({
            text: "-	Une ??tude de march?? par segmentation ad??quate aux actifs sur l???usage actuel et futur ;",
                bold: false,
            color: "000000",
            break: 1
            }),
            new TextRun({
                 text:"-	Analyse des offres comparables sur le march?? et des transactions r??centes effectu??es dans le p??rim??tre de l???actif ;",
                    bold: false,
                color: "000000",
                break: 1
                }),
            new TextRun({
                 text:"-	Valorisation des actifs en application de la m??thode par comparable et par capitalisation de revenu pour la d??finition de la valeur v??nale et marchande ;",
                       bold: false,
                   color: "000000",
                   break: 1
                   }),
        ],
          spacing: {
            after: 300,
            before : 20
        }
        }),
        clientTable,
        table8,
        new Paragraph({
            alignment: AlignmentType.CENTER,
            heading: HeadingLevel.TITLE,
            children: [
                new TextRun({
              text: "PRESENTATION DU BIEN ",
              bold: true,
              italics: true,
              color: '000000',
              allCaps: true,
              size: 50,
              break: 1,
              
    
                }),
            ],
            spacing: {
                before: 2500,
                after: 9000,
            },
            }),
            new Paragraph({
                  
                text:"         ",
                
                heading: HeadingLevel.TITLE,
               
               
                
                shading: {
                    type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                    color: "4682B4",
                    fill: "4682B4",
                },
                spacing: {
                    after: 200
                }
            }),
    new Paragraph({
                text: "\t2-1.	DEFINITION  DU BIEN	 	",
              
                heading: HeadingLevel.HEADING_1,
               
                spacing: {
                    after: 200
                }
            }),
            new Paragraph({
                text: "\t2-2.	VISITE DU BIEN	 	",
              
                heading: HeadingLevel.HEADING_1,
               
                spacing: {
                    after: 200
                }
            }),
            new Paragraph({
                text: "\t2-3.	DESCRIPTION  DU BIEN	 	",
              
                heading: HeadingLevel.HEADING_1,
               
                spacing: {
                    after: 1900
                }
            }),

            new Paragraph({
                children: [
                    new TextRun({
                        text : "Le bien expertis?? est "+ this.nature("ter", "bur", "ime", "com")+this.onSubmit('adresse') +", dans immeuble de prestige d???une chronologie r??cente. La surface exploitable de l???actif est de "+this.onSubmit('surface')+" m2 conform??ment au certificat de propri??t??. L???actif expertis?? se trouve dans une copropri??t??, comme la majorit?? des immeubles qui longent ladite avenue, le b??timent en question dispose abrite un rez-de-chauss??e commercial. Par ailleurs, il est n??cessaire de souligner que l???actif est situ?? au niveau du "+this.onSubmit('etage')+" ??tage sur un ensemble de type R+.",
                  bold: false,
                  color: "000000",
                  break: 2
              }),
            ],
              spacing: {
                after: 300,
                before : 20
            }
            }),
            new Paragraph({children: [
                new TextRun({
                    text:"        2-1   D??finition DU BIEN    ",
                    bold: true,
              color: "000000",
            }),],
            heading: HeadingLevel.HEADING_1,
            shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                color: "4682B4",
                fill: "4682B4",},
            spacing: {after: 200,
                before : 20}}),
                new Paragraph({children: [
                    new TextRun({
                        text: this.onSubmit1('bon'),
                        bold: false,
                  color: "000000",
                }),],
                spacing: {after: 200,
                    before : 20}
            }),
            new Paragraph({children: [
                new TextRun({
                    text: "Composition du bien : " + this.nature("ter", "bur", "ime", "com"),
                    bold: true,
              color: "000000",
              break: 1
            }),
            new TextRun({
                text: "Surface Totale : " + this.onSubmit('surface'),
                bold: true,
          color: "000000",
        }),
        ],
            spacing: {after: 200,
                before : 20}
        }),
        new Paragraph({children: [
            new TextRun({
                text:"        Atout du Bien    ",
                bold: true,
          color: "000000",
        }),],
        heading: HeadingLevel.HEADING_1,
        shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
            color: "4682B4",
            fill: "4682B4",},
        spacing: {after: 200,
            before : 20}}),
            new Paragraph({children: [
                
            new TextRun({
                text:  this.onSubmit1('atout'),
                bold: true,
          color: "000000",
        }),
        ],
            spacing: {after: 200,
                before : 20}
        }),

        new Paragraph({children: [
            new TextRun({
                text:"        Analyse de l'environement    ",
                bold: true,
          color: "000000",
        }),],
        heading: HeadingLevel.HEADING_1,
        shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
            color: "4682B4",
            fill: "4682B4",},
        spacing: {after: 200,
            before : 20}}),
            new Paragraph({children: [
                
            new TextRun({
                text:  this.onSubmit1('analyse'),
                bold: true,
          color: "000000",
        }),
        ],
            spacing: {after: 200,
                before : 20}
        }),
        new Paragraph({children: [
                
            new TextRun({
                text:  ">	 Description de l???environnement local et du quartier :",
                bold: true,
          color: "000000",
        }),
        ],
            spacing: {after: 200,
                before : 20}
        }),
        situationTable,
        new Paragraph({children: [
                
            new TextRun({
                text:  ">	 ??quipements & Services ?? proximit?? :",
                bold: true,
          color: "000000",
        }),
        ],
            spacing: {after: 200,
                before : 20}
        }),
                proxTable,
                new Paragraph({children: [
                    new TextRun({
                        text:"        2-2 visite de bien    ",
                        bold: true,
                  color: "000000",
                }),],
                heading: HeadingLevel.HEADING_1,
                shading: {type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                    color: "4682B4",
                    fill: "4682B4",},
                spacing: {after: 200,
                    before : 20}}),
                    new Paragraph({children: [
                
                        new TextRun({
                            text:  ">La visite du site a eu lieu le "+this.onSubmit('date')+" et a permis de contr??ler et de confirmer la nature de l???environnement par rapport aux documents urbanistiques qui nous ont ??t?? fournis.",
                            bold: false,
                      color: "000000",
                    }),
                    ],
                        spacing: {after: 200,
                            before : 20}
                    }),

                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        heading: HeadingLevel.TITLE,
                        children: [
                            new TextRun({
                          text: "Reportage photographique ",
                          bold: true,
                          italics: true,
                          color: '000000',
                          allCaps: true,
                          size: 50,
                          break: 1,
                          
                
                            }),
                        ],
                        spacing: {
                            before: 2500,
                            after: 9000,
                        },
                        }),
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            heading: HeadingLevel.TITLE,
                            children: [
                                new TextRun({
                              text: "EVALUATION DU BIEN ",
                              bold: true,
                              italics: true,
                              color: '000000',
                              allCaps: true,
                              size: 50,
                              break: 1,
                              
                    
                                }),
                            ],
                            spacing: {
                                before: 2500,
                                after: 9000,
                            },
                            }),
                            new Paragraph({
                                text: "\t4-1.	ETAT APPARENT DU BIEN	",
                              
                                heading: HeadingLevel.HEADING_1,
                               
                                spacing: {
                                    after: 200
                                }
                            }),
                            new Paragraph({
                                text: "\t4-2.	SITUATION FONCIERE ??? JURIDIQUE	 ",
                              
                                heading: HeadingLevel.HEADING_1,
                               
                                spacing: {
                                    after: 200
                                }
                            }),
                            new Paragraph({
                                text: "\t4-3.	ETUDE DE MARCHE	",
                              
                                heading: HeadingLevel.HEADING_1,
                               
                                spacing: {
                                    after: 1900
                                }
                            }),
                            new Paragraph({children: [
                
                                new TextRun({
                                    text: "Dans cette partie de l?????tude, il s???av??re pertinent de s???attarder sur les points essentiels qui pourront donner une vision plus explicite sur l???actif expertis??. Il s???agit notamment du volet technique juridique, et commercial. Ces trois volets viennent clarifier l???essence m??me de cette expertise par l???effectivit?? de l?????tat apparent du bien et son ??tat des lieux ainsi que l?????tat du march?? mitoyen.Dans le cadre de cette ??tude et en premier lieu, pour l?????tude de la partie technique, nous serons amen??s ?? analyser les diff??rents aspects relatifs ?? l?????tat apparent du second ??uvre, des finitions et notamment l?????tat des parties communes de la r??sidence. Un syst??me de notation sera mis en place dans le but de d??termi ner les diff??rentes contraintes. Cette notation sera ??labor??e suite aux constats g??n??raux que nous aurons observ??s lors de la visite sur le terrain.",
                                    bold: false,
                              color: "000000",
                            }),
                            ],
                                spacing: {after: 200,
                                    before : 20}
                            }),

                            new Paragraph({children: [
                
                                new TextRun({
                                    text: "     4-1  Etat apparent du bien ",
                                    bold: true,
                              color: "000000",
                            }),
                            ],
                                spacing: {after: 200,
                                    before : 20}
                            }),
                            new Paragraph({children: [
                
                                new TextRun({
                                    text:"L?????tude technique tend ?? identifier l???aspect et les contraintes ayant un impact direct sur l???actif expertis??. Ce volet d???analyse permet de faire un constat tout en examinant l?????tat apparent et l?????tat technique du bien sur leurs ??tats apparents sous r??serve d???une ??tude technique approfondie. Un syst??me de notation est adopt?? en vue d???avoir une vision claire sur l?????tat du site. Dans ce volet d???analyse, notre objectif est de d??montrer si le site b??n??ficie d???une accessibilit??, de branche- ments, de risques de pollutions et autres risques pouvant alt??rer le site et par cons??quent compromettre sa valorisation",
                                    bold: false,
                              color: "000000",
                            }),
                            ],
                                spacing: {after: 200,
                                    before : 20}
                            }),
                            table9,
                            new Paragraph({children: [
                
                                new TextRun({
                                    text: "     4-2.	SITUATION FONCIERE ??? JURIDIQUE",
                                    bold: true,
                              color: "000000",
                            }),
                            ],
                                spacing: {after: 200,
                                    before : 20}
                            }),
                            new Paragraph({children: [
                
                                new TextRun({
                                    text: "Les informations juridiques consign??es sur le certificat de propri??t?? communiqu??, se synth??tisent de la mani??re suivante :",
                                    bold: false,
                              color: "000000",
                            }),
                            ],
                                spacing: {after: 200,
                                    before : 20}
                            }),
                            juridiqueTable,
                            new Paragraph({children: [
                
                                new TextRun({
                                    text: "     4-3.	Etude de march??",
                                    bold: true,
                              color: "000000",
                            }),
                            ],
                                spacing: {after: 200,
                                    before : 20}
                            }),
                            new Paragraph({children: [
                
                                new TextRun({
                                    text: "Dans cet axe, notre ??tude de march?? nous a amen??e ?? faire est un travail de collecte des donn??es immobi- li??res et d???analyse d???informations dans le but d???identifier les caract??ristiques du march?? actuel dans l???immo- bilier et notamment ?? des biens comparables ?? l???actif expertis??. Le nombre d???offres de bureaux sur le march?? de l???Avenue Omar ibnou AL khattab est tr??s reduit. En effet, ce march?? propose essentiellement des appartements  de surfaces interm??diaires et subit une grosse demande qui cause la raret?? d???offre de bureaux dans le boulevard. Malgr?? un niveau de prix croissant,  l???avenue Omar Ibn Al Khattab ?? privil??gier lors d???un investissement locatif, notamment si vous souhaitez optimiser la valorisation de votre patrimoine sur le long terme. En effet, l???immobilier locatif d??tient de nombreux atouts ??conomiques, touristiques et culturels, qui augmentent le prix dss locaux, mais qui font du boulevard une valeur s??re pour les investisseurs vu la raret?? des locaux a usage profession??l dans le quartier hormis deux ou trois immeuble a usage de bureaux dans un rayon de 500m a 1KM. Notre ??tude de march?? prend en compte plusieurs crit??res notamment les statistiques en termes de popula- tion dans la zone, le nombre d???offres pr??sentes ainsi que la demande et les dur??es d?????coulements des actifs mises ?? la vente. La valeur homog??n??is??e moyenne d??gag??e de cette ??tude de march?? a ??t?? retenue pour pour les calculs de la pr??sente expertise. L???analyse du march?? concurrentiel met en ??vidence des prix moyens de vente, compris entre 9500 MAD/ m?? et 11.500 MAD/m?? sur des actifs de type appartement dans la zone.",
                                    bold: true,
                              color: "000000",
                            }),
                            ],
                                spacing: {after: 200,
                                    before : 20}
                            }),
                            table10,



          ],
      }],
  });
  
  // Used to export the file into a .docx file
  // Packer.toBuffer(doc).then((buffer) => {
  //     fs.writeFileSync("My Document.docx", buffer);
  //     console.log("Document created successfully");
  
  // });
  Packer.toBlob(doc).then(blob => {
      console.log(blob);
      saveAs(blob, "Rapport.docx");
      console.log("Document created successfully");
    });
      
    }
}
