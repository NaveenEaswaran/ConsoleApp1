using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IntakeBase = Xom.Gci.Addin.LvMake.Intake.Base;
using IIntakeBase = Xom.Gci.Addin.LvMake.IIntake;
namespace ConsoleApp1
{
   public  class NewClass: IIntakeBase.IIntakeLIMSDataFetch  
    {

        LIMSDataModel FetchAndLoadMasterData(LvHelper lvHelper, string absoluteURL, string requestItemID);
        void FetchAndLoadSheetData(LvHelper LvHelper, LIMSDataModel limsModel, IIntakeValidate validate, IIntakeContext context, IIntakeIngredient ingredient, IIntakeFormulation formulation, IIntakeBlends blends, IIntakeProcessVariables processVariables, IIntakeActualIng actualIng, IIntakeBatchContainer batchContainer, IIntakeInventory inventory, IIntakeRibbon ribbon, ConditionalFormattingHelper cfHelper, ColumnWidthHelper widthHelper, IIntakeTreatedIngredient treatedIngredient, IIntakeTreatedBatches treatedBatches, IIntakeTests tests, IIntakeContainerProcessVariables containerProcessVariables, IIntakeReviewComplete reviewComplete, bool isSimpleIntake = true);
        void FetchAndLoadColorData(LIMSDataModel limsDataModel);

    }
}
