// Programme cod� en JavaScript permettant de g�rer l'entit� Bien

if (typeof (PackageBien) == "undefined")
{ PackageBien = {}; }
PackageBien.Bien = {
  onBienLoad: function () {//PackageBien.Bien.onBienLoad
    setSubmitMode("resi_tvautiliseetype", "always");
    setSubmitMode("resi_prixtvaamenagee", "always");
    setSubmitMode("resi_montantadv", "always");
    setSubmitMode("resi_prixcontrat", "always");
  },
}