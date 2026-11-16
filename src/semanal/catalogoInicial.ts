// Catálogo inicial (desde la planilla Pedidos_Semanales 2026): especies, mercados y
// proveedores con las especies que trabajan. Se carga una sola vez desde la sección
// semanal si la base aún no tiene catálogo. Después se edita en Ajustes.
import type { Catalogo } from "./datos";

export const CATALOGO_INICIAL: Catalogo = {
  "especies": [
    "Cerezas",
    "Nectarines",
    "Nectarines 16 kg",
    "Arándanos",
    "Cebolla",
    "Duraznos",
    "Damasco",
    "Naranja",
    "Ciruela",
    "Ciruela 16 kg",
    "Clementinas/ Mandarinas",
    "Palta",
    "Kiwi",
    "Kiwi 16 kg",
    "Kiwi Gold",
    "Limón",
    "Manzana Opal",
    "Manzana Gala",
    "Manzana Gala Granel",
    "Manzana Granny",
    "Manzana Granny Granel",
    "Manzana Roja",
    "Manzana Cripps Pink",
    "Manzana Pink Lady",
    "Manzana Fuji",
    "Manzana Braeburn",
    "Membrillo 7,2 kg",
    "Pomelo",
    "Peras Forelle",
    "Peras Flamingo",
    "Peras Asiáticas",
    "Peras Bosc",
    "Peras Bartlett",
    "Peras Packhams",
    "Peras Coscia",
    "Peras Williams",
    "Granada",
    "Red Globe",
    "Uva Negra",
    "Uva Roja",
    "Uva Verde",
    "Peras Carmen",
    "Peras D'Anjou",
    "Peras D'Anjou Red"
  ],
  "mercados": [
    {
      "code": "EC",
      "nombre": "Ecuador",
      "empresa": "Frutam"
    },
    {
      "code": "COL",
      "nombre": "Colombia",
      "empresa": "Frutam"
    },
    {
      "code": "PAN",
      "nombre": "Panamá",
      "empresa": "Frutam"
    },
    {
      "code": "CR",
      "nombre": "Costa Rica",
      "empresa": "Frutam"
    },
    {
      "code": "SV",
      "nombre": "El Salvador",
      "empresa": "Frutam"
    },
    {
      "code": "GT",
      "nombre": "Guatemala",
      "empresa": "Frutam"
    },
    {
      "code": "MX",
      "nombre": "México",
      "empresa": "Frutam"
    },
    {
      "code": "HOL",
      "nombre": "Holanda",
      "empresa": "Polar"
    }
  ],
  "proveedores": [
    {
      "nombre": "Agricom",
      "especies": [
        "Cerezas",
        "Nectarines",
        "Arándanos",
        "Duraznos",
        "Naranja",
        "Ciruela",
        "Clementinas/ Mandarinas",
        "Palta",
        "Limón",
        "Manzana Gala",
        "Pomelo",
        "Granada",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde"
      ]
    },
    {
      "nombre": "Agrisouth",
      "especies": [
        "Manzana Gala",
        "Manzana Roja"
      ]
    },
    {
      "nombre": "Agrofruta",
      "especies": [
        "Kiwi",
        "Manzana Gala",
        "Manzana Pink Lady",
        "Manzana Fuji",
        "Peras Asiáticas",
        "Peras Bosc",
        "Peras Packhams",
        "Peras Coscia",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde"
      ]
    },
    {
      "nombre": "Agrosan",
      "especies": [
        "Cerezas",
        "Nectarines",
        "Arándanos",
        "Duraznos",
        "Ciruela",
        "Palta",
        "Manzana Gala",
        "Peras Bartlett",
        "Peras Coscia",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde",
        "Peras Carmen"
      ]
    },
    {
      "nombre": "Agua Santa",
      "especies": [
        "Cerezas",
        "Nectarines",
        "Duraznos",
        "Ciruela",
        "Kiwi",
        "Peras Bartlett",
        "Peras Packhams",
        "Peras Coscia",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde"
      ]
    },
    {
      "nombre": "Almahue",
      "especies": []
    },
    {
      "nombre": "Andes Bhumi",
      "especies": [
        "Kiwi",
        "Manzana Granny",
        "Peras Forelle"
      ]
    },
    {
      "nombre": "Atacama",
      "especies": [
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde"
      ]
    },
    {
      "nombre": "Baika",
      "especies": [
        "Naranja",
        "Clementinas/ Mandarinas",
        "Palta",
        "Limón"
      ]
    },
    {
      "nombre": "C&L",
      "especies": [
        "Cerezas",
        "Nectarines",
        "Duraznos",
        "Ciruela"
      ]
    },
    {
      "nombre": "Cerasus",
      "especies": [
        "Cerezas",
        "Arándanos"
      ]
    },
    {
      "nombre": "CFN",
      "especies": [
        "Cerezas",
        "Nectarines",
        "Arándanos",
        "Duraznos",
        "Ciruela",
        "Clementinas/ Mandarinas",
        "Palta",
        "Manzana Gala",
        "Peras Bartlett",
        "Peras Coscia",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde",
        "Peras Carmen"
      ]
    },
    {
      "nombre": "Copefrut",
      "especies": [
        "Cerezas",
        "Nectarines",
        "Duraznos",
        "Ciruela",
        "Kiwi",
        "Manzana Gala",
        "Manzana Granny",
        "Manzana Fuji",
        "Peras Bartlett",
        "Peras Coscia",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde",
        "Peras Carmen"
      ]
    },
    {
      "nombre": "Crispagold",
      "especies": [
        "Naranja",
        "Clementinas/ Mandarinas",
        "Limón"
      ]
    },
    {
      "nombre": "Del Monte",
      "especies": [
        "Cerezas",
        "Nectarines",
        "Arándanos",
        "Duraznos",
        "Ciruela",
        "Palta",
        "Manzana Gala",
        "Peras Bartlett",
        "Peras Coscia",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde",
        "Peras Carmen"
      ]
    },
    {
      "nombre": "Dole",
      "especies": [
        "Ciruela",
        "Clementinas/ Mandarinas",
        "Manzana Gala",
        "Manzana Granny",
        "Manzana Roja",
        "Manzana Pink Lady",
        "Peras Forelle",
        "Peras Bosc",
        "Peras Bartlett",
        "Peras Packhams",
        "Peras Coscia"
      ]
    },
    {
      "nombre": "Dole (ARG)",
      "especies": [
        "Peras Bosc"
      ]
    },
    {
      "nombre": "Dosal",
      "especies": [
        "Cerezas",
        "Nectarines",
        "Duraznos",
        "Ciruela",
        "Manzana Gala",
        "Manzana Gala Granel",
        "Manzana Granny",
        "Manzana Roja",
        "Manzana Cripps Pink",
        "Peras Bartlett",
        "Peras Coscia",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde",
        "Peras Carmen"
      ]
    },
    {
      "nombre": "Eastern Island",
      "especies": [
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde"
      ]
    },
    {
      "nombre": "El Camino",
      "especies": [
        "Nectarines",
        "Duraznos",
        "Ciruela",
        "Palta",
        "Kiwi",
        "Manzana Gala",
        "Manzana Granny",
        "Peras Bartlett",
        "Peras Coscia",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde",
        "Peras Carmen"
      ]
    },
    {
      "nombre": "El Parque",
      "especies": [
        "Naranja",
        "Clementinas/ Mandarinas",
        "Palta",
        "Limón",
        "Pomelo"
      ]
    },
    {
      "nombre": "Erfrut",
      "especies": [
        "Manzana Gala"
      ]
    },
    {
      "nombre": "Espejo",
      "especies": [
        "Cerezas",
        "Nectarines",
        "Duraznos",
        "Ciruela",
        "Manzana Gala",
        "Manzana Granny",
        "Manzana Roja",
        "Manzana Cripps Pink",
        "Peras Coscia",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde",
        "Peras Carmen"
      ]
    },
    {
      "nombre": "Exportal",
      "especies": [
        "Cerezas",
        "Nectarines",
        "Duraznos",
        "Ciruela",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde"
      ]
    },
    {
      "nombre": "Frusan",
      "especies": [
        "Cerezas",
        "Nectarines",
        "Duraznos",
        "Kiwi",
        "Manzana Gala",
        "Manzana Granny",
        "Manzana Pink Lady",
        "Manzana Braeburn",
        "Membrillo 7,2 kg",
        "Peras Forelle",
        "Peras Asiáticas",
        "Peras Packhams",
        "Peras Coscia",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde",
        "Peras Carmen"
      ]
    },
    {
      "nombre": "Fruta Austral",
      "especies": [
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde"
      ]
    },
    {
      "nombre": "Frutasol",
      "especies": [
        "Manzana Gala",
        "Manzana Granny",
        "Manzana Roja",
        "Manzana Pink Lady",
        "Manzana Fuji"
      ]
    },
    {
      "nombre": "Geofrut",
      "especies": [
        "Nectarines",
        "Duraznos",
        "Ciruela",
        "Kiwi",
        "Manzana Gala",
        "Peras Packhams",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde"
      ]
    },
    {
      "nombre": "Gesex",
      "especies": [
        "Nectarines",
        "Duraznos",
        "Ciruela",
        "Kiwi",
        "Red Globe",
        "Uva Negra",
        "Uva Verde"
      ]
    },
    {
      "nombre": "Gioia",
      "especies": [
        "Nectarines",
        "Naranja",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde"
      ]
    },
    {
      "nombre": "Greenex",
      "especies": [
        "Manzana Gala",
        "Manzana Granny"
      ]
    },
    {
      "nombre": "Greenvic",
      "especies": [
        "Cerezas",
        "Nectarines",
        "Arándanos",
        "Duraznos",
        "Ciruela",
        "Kiwi",
        "Manzana Gala",
        "Manzana Granny",
        "Manzana Roja",
        "Manzana Pink Lady",
        "Manzana Fuji",
        "Peras Forelle",
        "Peras Packhams",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde"
      ]
    },
    {
      "nombre": "Kleppe",
      "especies": [
        "Peras Packhams"
      ]
    },
    {
      "nombre": "Laberinto",
      "especies": [
        "Nectarines",
        "Duraznos",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde"
      ]
    },
    {
      "nombre": "LAFRUT",
      "especies": [
        "Cerezas",
        "Nectarines",
        "Arándanos",
        "Duraznos",
        "Ciruela",
        "Palta",
        "Manzana Gala",
        "Peras Forelle",
        "Peras Flamingo",
        "Peras Packhams",
        "Peras Coscia",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde",
        "Peras Carmen"
      ]
    },
    {
      "nombre": "Lapacho Azul (ARG)",
      "especies": [
        "Limón"
      ]
    },
    {
      "nombre": "Los Lirios",
      "especies": [
        "Manzana Gala",
        "Manzana Granny"
      ]
    },
    {
      "nombre": "Magna",
      "especies": [
        "Nectarines",
        "Duraznos",
        "Manzana Gala",
        "Manzana Gala Granel",
        "Manzana Granny",
        "Manzana Roja",
        "Peras Forelle",
        "Peras Flamingo",
        "Peras Bosc",
        "Peras Packhams",
        "Peras Coscia",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde"
      ]
    },
    {
      "nombre": "MDT",
      "especies": [
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde"
      ]
    },
    {
      "nombre": "Monte Alto",
      "especies": [
        "Cerezas",
        "Nectarines",
        "Duraznos",
        "Ciruela",
        "Peras Coscia",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde",
        "Peras D'Anjou"
      ]
    },
    {
      "nombre": "Moño Azul (AR)",
      "especies": [
        "Peras Bosc",
        "Peras Packhams",
        "Peras Coscia",
        "Peras D'Anjou",
        "Peras D'Anjou Red"
      ]
    },
    {
      "nombre": "PFI",
      "especies": [
        "Cerezas",
        "Kiwi",
        "Manzana Opal",
        "Manzana Gala",
        "Manzana Granny",
        "Red Globe",
        "Uva Verde"
      ]
    },
    {
      "nombre": "Pocuro",
      "especies": [
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde"
      ]
    },
    {
      "nombre": "Prize",
      "especies": [
        "Cerezas",
        "Nectarines",
        "Arándanos",
        "Duraznos",
        "Ciruela",
        "Clementinas/ Mandarinas",
        "Palta",
        "Manzana Gala",
        "Peras Bartlett",
        "Peras Coscia",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde",
        "Peras Carmen"
      ]
    },
    {
      "nombre": "Propal",
      "especies": [
        "Palta",
        "Kiwi",
        "Kiwi Gold"
      ]
    },
    {
      "nombre": "PROVEX",
      "especies": [
        "Red Globe",
        "Uva Verde"
      ]
    },
    {
      "nombre": "Quelén",
      "especies": [
        "Naranja"
      ]
    },
    {
      "nombre": "Ranco",
      "especies": [
        "Cerezas"
      ]
    },
    {
      "nombre": "Rio King",
      "especies": [
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde"
      ]
    },
    {
      "nombre": "Rucaray",
      "especies": [
        "Manzana Gala",
        "Manzana Granny"
      ]
    },
    {
      "nombre": "Río Blanco",
      "especies": [
        "Naranja",
        "Granada",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde"
      ]
    },
    {
      "nombre": "San Lorenzo",
      "especies": [
        "Clementinas/ Mandarinas",
        "Red Globe"
      ]
    },
    {
      "nombre": "San Luis de Yaquil",
      "especies": [
        "Cerezas",
        "Nectarines",
        "Duraznos",
        "Ciruela",
        "Palta",
        "Manzana Gala",
        "Manzana Granny",
        "Peras Bartlett",
        "Peras Coscia",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde",
        "Peras Carmen"
      ]
    },
    {
      "nombre": "Schmidt",
      "especies": [
        "Cerezas",
        "Naranja",
        "Clementinas/ Mandarinas",
        "Palta",
        "Limón"
      ]
    },
    {
      "nombre": "SM Export",
      "especies": [
        "Nectarines",
        "Duraznos",
        "Ciruela",
        "Ciruela 16 kg",
        "Palta",
        "Kiwi",
        "Manzana Gala",
        "Manzana Granny",
        "Peras Packhams",
        "Peras Coscia",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde",
        "Peras Carmen"
      ]
    },
    {
      "nombre": "Sofruco",
      "especies": [
        "Cerezas",
        "Nectarines",
        "Arándanos",
        "Duraznos",
        "Naranja",
        "Ciruela",
        "Clementinas/ Mandarinas",
        "Palta",
        "Manzana Gala",
        "Peras Bartlett",
        "Peras Packhams",
        "Peras Coscia",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde",
        "Peras Carmen"
      ]
    },
    {
      "nombre": "Subsole",
      "especies": [
        "Cerezas",
        "Naranja",
        "Clementinas/ Mandarinas",
        "Palta",
        "Kiwi",
        "Limón",
        "Granada",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde"
      ]
    },
    {
      "nombre": "Total Fruit",
      "especies": [
        "Cerezas",
        "Nectarines",
        "Arándanos",
        "Duraznos",
        "Ciruela",
        "Palta",
        "Manzana Gala"
      ]
    },
    {
      "nombre": "Triofrut",
      "especies": [
        "Cerezas",
        "Nectarines",
        "Duraznos",
        "Ciruela",
        "Kiwi",
        "Peras Forelle",
        "Peras Packhams"
      ]
    },
    {
      "nombre": "Verfrut",
      "especies": [
        "Nectarines",
        "Nectarines 16 kg",
        "Duraznos",
        "Ciruela",
        "Ciruela 16 kg",
        "Kiwi",
        "Manzana Gala",
        "Manzana Granny",
        "Manzana Cripps Pink",
        "Manzana Fuji",
        "Peras Asiáticas",
        "Peras Bartlett",
        "Peras Coscia",
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde"
      ]
    },
    {
      "nombre": "WHM",
      "especies": [
        "Red Globe",
        "Uva Negra",
        "Uva Roja",
        "Uva Verde"
      ]
    },
    {
      "nombre": "Yanten",
      "especies": [
        "Manzana Gala",
        "Manzana Granny"
      ]
    }
  ]
};
