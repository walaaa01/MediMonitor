#include <stdio.h>
#include <stdlib.h>

struct noeud{
    int cle;
    struct noeud suivant;
};

void affCond(struct noeud *p, unsigned(*oper)(struct noeud *)){
    while((*oper)(p)){
        printf("%d\n",p->cle);
        p=p->suivant;
    }
}

unsigned condition(struct noeud *p){
    return p->cle >=5;
}

void main(){
    struct noeud *tst;
    creer_liste(tst);
    affCond(tst,condition);
}
