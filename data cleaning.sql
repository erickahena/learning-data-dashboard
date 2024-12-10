/****** Script for SelectTopNRows command from SSMS  ******/
SELECT TOP (1000) [UniqueID ]
      ,[ParcelID]
      ,[LandUse]
      ,[PropertyAddress]
      ,[SaleDate]
      ,[SalePrice]
      ,[LegalReference]
      ,[SoldAsVacant]
      ,[OwnerName]
      ,[OwnerAddress]
      ,[Acreage]
      ,[TaxDistrict]
      ,[LandValue]
      ,[BuildingValue]
      ,[TotalValue]
      ,[YearBuilt]
      ,[Bedrooms]
      ,[FullBath]
      ,[HalfBath]
      ,[SaleDateConverted]
  FROM [Portfolio].[dbo].[Sheet1$]

  -- standardize date format
  select SaleDateConverted, CONVERT(Date,SaleDate)
  From Portfolio.dbo.Sheet1$

  update Sheet1$
  set SaleDate = CONVERT(date,SaleDate)

  alter table dbo.Sheet1$
  add SaleDateConverted Date;
  
  update dbo.Sheet1$
  set SaleDate = CONVERT(date,SaleDate)

  --populate property address data

  select*
  from dbo.Sheet1$
  --where propertyaddress is null
  order by ParcelID

  select a.ParcelID, a.PropertyAddress, b.ParcelID, b.PropertyAddress, ISNULL(a.Propertyaddress,b.PropertyAddress)
  from Portfolio.dbo.Sheet1$ a
  join Portfolio.dbo.Sheet1$ b
  on a.ParcelID = b.ParcelID
  and a.[UniqueID ] <> b.[UniqueID ]
  where a.PropertyAddress is null

  update a
  set PropertyAddress = ISNULL(a.Propertyaddress,b.PropertyAddress)
  from Portfolio.dbo.Sheet1$ a
  join Portfolio.dbo.Sheet1$ b
  on a.ParcelID = b.ParcelID
  and a.[UniqueID ] <> b.[UniqueID ]
   where a.PropertyAddress is null

   --breaking out address into individual columns (address, city, state)

   select PropertyAddress
   from dbo.Sheet1$
  --where propertyaddress is null
  --order by ParcelID

  select
  SUBSTRING (PropertyAddress, 1, CHARINDEX(',',PropertyAddress)-1) as Address,
  SUBSTRING (PropertyAddress,charindex(',',PropertyAddress)+1, len(PropertyAddress)) as Address
  from dbo.Sheet1$

  alter table Sheet1$
  add PropertySplitAddress nvarchar(255)

  update Sheet1$
  set PropertySplitAddress =  SUBSTRING (PropertyAddress, 1, CHARINDEX(',',PropertyAddress)-1)

  alter table Sheet1$
  add PropertySplitCity nvarchar(255)

  update Sheet1$
  set PropertySplitCity = SUBSTRING (PropertyAddress,charindex(',',PropertyAddress)+1, len(PropertyAddress))
  
  select*
  from dbo.Sheet1$

  select
  PARSENAME(replace(OwnerAddress,',','.'),3)
  ,PARSENAME(replace(OwnerAddress,',','.'),2)
  ,PARSENAME(replace(OwnerAddress,',','.'),1)
  from dbo.Sheet1$

  alter table Sheet1$
  add OwnerSplitAddress nvarchar(255)

  update Sheet1$
  set OwnerSplitAddress = PARSENAME(replace(OwnerAddress,',','.'),3)

  alter table Sheet1$
  add OwnerSplitCity nvarchar(255)

  update Sheet1$
  set OwnerSplitCity = PARSENAME(replace(OwnerAddress,',','.'),2)

  alter table Sheet1$
  add OwnerSplitState nvarchar(255)

  update Sheet1$
  set OwnerSplitState = PARSENAME(replace(OwnerAddress,',','.'),1)
  
  select*
  from dbo.Sheet1$

  --change Y and N and No in "Sold as Vacant" field

  Select distinct(SoldAsVacant), Count(SoldASVacant)
  From dbo.Sheet1$
  Group by SoldAsVacant
  order by 2

  select distinct(UniqueID), count(UniqueID)
  from dbo.Sheet1$
  group by [UniqueID ]

  select SoldAsVacant
  , case when SoldAsVacant = 'Y' then 'Yes'
  when SoldAsVacant = 'N' then 'No'
  else SoldAsVacant
  end
  From dbo.Sheet1$

  update dbo.Sheet1$
  set SoldAsVacant = case when SoldAsVacant = 'Y' then 'Yes'
  when SoldAsVacant = 'N' then 'No'
  else SoldAsVacant
  end

  --remove duplicates

  WITH RowNumCTE AS(
  select*,
  ROW_NUMBER() OVER(
  PARTITION BY PARCELID,
				Propertyaddress,
				saleprice,
				saledate,
				legalreference
				ORDER BY UniqueID) row_num

  from dbo.Sheet1$
  --order by ParcelID
  )
  select*
  from RowNumCTE
  where row_num > 1
  order by PropertyAddress

  --delete unused columns

  select*
  from dbo.Sheet1$

  alter table dbo.Sheet1$
  drop column OwnerAddress, TaxDistrict, PropertyAddress
  
  alter table dbo.Sheet1$
  drop column SaleDate
