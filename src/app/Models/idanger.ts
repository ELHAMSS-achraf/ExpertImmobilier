import { Logged } from "./logged";

export interface IDanger {
    id: number,
    user: Logged,
    type:string,
    date: Date,
    commentaire : string,
    longitude:number,
    latitude: number,
}
