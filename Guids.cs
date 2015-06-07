// Guids.cs
// MUST match guids.h
using System;

namespace Utkarsh.HierarchyDemo
{
    static class GuidList
    {
        public const string guidHierarchyDemoPkgString = "2d96c484-eb03-433b-b279-1ee39fa38cd0";
        public const string guidHierarchyDemoCmdSetString = "33563ee4-86e3-4b21-b04d-e9f78370d345";
        public const string guidToolWindowPersistanceString = "c3fc9d92-4a03-46a4-abee-047529674048";

        public static readonly Guid guidHierarchyDemoCmdSet = new Guid(guidHierarchyDemoCmdSetString);


        public const string guidHierAnarchyPersistanceString = "e143afdc-31ea-4bde-8392-13d02392cfbc";
        public static readonly Guid guidHierAnarchyPersistance = new Guid(guidHierAnarchyPersistanceString);
    };
}